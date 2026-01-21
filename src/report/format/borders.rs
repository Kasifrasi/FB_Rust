//! Border-System mit Merge-Awareness
//!
//! Dieses Modul stellt ein range-basiertes Border-System bereit, das:
//! - Borders auf Bereiche anwendet (nicht einzelne Zellen)
//! - Merged Cells erkennt und respektiert
//! - Explizite Fehler wirft bei ungültigen Border-Positionen
//!
//! # Workflow
//!
//! ```ignore
//! // 1. ZUERST: Merged Cells registrieren
//! let mut merged_cells = MergedCellRegistry::new();
//! merged_cells.register_merge(0, 4, 1, 4);  // E1:E2
//!
//! // 2. DANN: BorderManager erstellen und Merges übergeben
//! let mut borders = BorderManager::new();
//! borders.set_merged_cells(merged_cells);
//!
//! // 3. DANACH: Borders definieren
//! borders.add_range(0, 4, 1, 4, BorderSpec::all_medium());
//!
//! // 4. ZULETZT: Borders anwenden
//! for row in 0..10 {
//!     for col in 0..10 {
//!         match borders.get_border_for_cell(row, col) {
//!             Ok(Some(border_fmt)) => {
//!                 // Border via merge() anwenden
//!             }
//!             Ok(None) => {
//!                 // Keine Border
//!             }
//!             Err(e) => {
//!                 eprintln!("Border error: {:?}", e);
//!             }
//!         }
//!     }
//! }
//! ```

use rust_xlsxwriter::{Format, FormatBorder};
use std::fmt;

/// Border-Spezifikation für einen Bereich
#[derive(Debug, Clone)]
pub struct BorderSpec {
    pub top: Option<FormatBorder>,
    pub bottom: Option<FormatBorder>,
    pub left: Option<FormatBorder>,
    pub right: Option<FormatBorder>,
}

impl BorderSpec {
    /// Alle vier Seiten mit dünner Border
    pub fn all_thin() -> Self {
        Self {
            top: Some(FormatBorder::Thin),
            bottom: Some(FormatBorder::Thin),
            left: Some(FormatBorder::Thin),
            right: Some(FormatBorder::Thin),
        }
    }

    /// Alle vier Seiten mit mittlerer Border
    pub fn all_medium() -> Self {
        Self {
            top: Some(FormatBorder::Medium),
            bottom: Some(FormatBorder::Medium),
            left: Some(FormatBorder::Medium),
            right: Some(FormatBorder::Medium),
        }
    }

    /// Alle vier Seiten mit dicker Border
    pub fn all_thick() -> Self {
        Self {
            top: Some(FormatBorder::Thick),
            bottom: Some(FormatBorder::Thick),
            left: Some(FormatBorder::Thick),
            right: Some(FormatBorder::Thick),
        }
    }

    /// Nur oben und unten (horizontal)
    pub fn top_bottom(border: FormatBorder) -> Self {
        Self {
            top: Some(border),
            bottom: Some(border),
            left: None,
            right: None,
        }
    }

    /// Nur links und rechts (vertikal)
    pub fn left_right(border: FormatBorder) -> Self {
        Self {
            top: None,
            bottom: None,
            left: Some(border),
            right: Some(border),
        }
    }

    /// Nur untere Border
    pub fn bottom_only(border: FormatBorder) -> Self {
        Self {
            top: None,
            bottom: Some(border),
            left: None,
            right: None,
        }
    }
}

/// Merged Cell Registry
///
/// WICHTIG: Alle Merged Cells müssen VOR Border-Anwendung registriert sein,
/// damit BorderManager keine Borders in der Mitte von Merged Cells setzt.
#[derive(Debug, Clone)]
pub struct MergedCellRegistry {
    merges: Vec<MergedRange>,
}

#[derive(Debug, Clone)]
struct MergedRange {
    start_row: u32,
    start_col: u16,
    end_row: u32,
    end_col: u16,
}

impl MergedCellRegistry {
    /// Erstellt eine neue, leere Registry
    pub fn new() -> Self {
        Self { merges: Vec::new() }
    }

    /// Registriert einen Merge-Bereich
    ///
    /// # Arguments
    ///
    /// * `start_row` - Erste Zeile (0-indexed)
    /// * `start_col` - Erste Spalte (0-indexed)
    /// * `end_row` - Letzte Zeile (0-indexed, inklusiv)
    /// * `end_col` - Letzte Spalte (0-indexed, inklusiv)
    pub fn register_merge(&mut self, start_row: u32, start_col: u16, end_row: u32, end_col: u16) {
        self.merges.push(MergedRange {
            start_row,
            start_col,
            end_row,
            end_col,
        });
    }

    /// Prüft, ob eine Zelle Teil einer Merged Cell ist (aber NICHT top-left)
    pub fn is_merged_non_topleft(&self, row: u32, col: u16) -> bool {
        for merge in &self.merges {
            if row >= merge.start_row
                && row <= merge.end_row
                && col >= merge.start_col
                && col <= merge.end_col
            {
                // Ist Teil des Merge, aber NICHT top-left?
                if row != merge.start_row || col != merge.start_col {
                    return true;
                }
            }
        }
        false
    }

    /// Gibt den Merge-Bereich zurück, wenn die Zelle die top-left einer Merge ist
    ///
    /// Returns: Some((start_row, start_col, end_row, end_col)) oder None
    pub fn get_merge_range(&self, row: u32, col: u16) -> Option<(u32, u16, u32, u16)> {
        for merge in &self.merges {
            if row == merge.start_row && col == merge.start_col {
                return Some((
                    merge.start_row,
                    merge.start_col,
                    merge.end_row,
                    merge.end_col,
                ));
            }
        }
        None
    }

    /// Gibt die Anzahl registrierter Merges zurück
    pub fn count(&self) -> usize {
        self.merges.len()
    }
}

impl Default for MergedCellRegistry {
    fn default() -> Self {
        Self::new()
    }
}

/// Border-Manager mit Merge-Awareness
///
/// Wendet Borders auf Bereiche an und respektiert dabei Merged Cells.
pub struct BorderManager {
    ranges: Vec<BorderRange>,
    merged_cells: Option<MergedCellRegistry>,
}

#[derive(Debug)]
struct BorderRange {
    start_row: u32,
    start_col: u16,
    end_row: u32,
    end_col: u16,
    spec: BorderSpec,
}

impl BorderManager {
    /// Erstellt einen neuen BorderManager
    pub fn new() -> Self {
        Self {
            ranges: Vec::new(),
            merged_cells: None,
        }
    }

    /// Setzt die Merged Cell Registry
    ///
    /// WICHTIG: MUSS aufgerufen werden bevor `get_border_for_cell()` verwendet wird!
    pub fn set_merged_cells(&mut self, registry: MergedCellRegistry) {
        self.merged_cells = Some(registry);
    }

    /// Registriert eine Border für einen Bereich
    ///
    /// # Arguments
    ///
    /// * `start_row` - Erste Zeile (0-indexed)
    /// * `start_col` - Erste Spalte (0-indexed)
    /// * `end_row` - Letzte Zeile (0-indexed, inklusiv)
    /// * `end_col` - Letzte Spalte (0-indexed, inklusiv)
    /// * `spec` - Border-Spezifikation
    pub fn add_range(
        &mut self,
        start_row: u32,
        start_col: u16,
        end_row: u32,
        end_col: u16,
        spec: BorderSpec,
    ) {
        self.ranges.push(BorderRange {
            start_row,
            start_col,
            end_row,
            end_col,
            spec,
        });
    }

    /// Gibt Border-Format für eine Zelle zurück (Merge-Aware!)
    ///
    /// # Errors
    ///
    /// - `BorderError::MergedCellsNotSet` wenn Merged Cell Registry nicht gesetzt wurde
    /// - `BorderError::BorderOnMergedNonTopLeft` wenn versucht wird, Border auf
    ///   non-topleft merged cell zu setzen
    ///
    /// # Returns
    ///
    /// - `Ok(Some(Format))` wenn Border für diese Zelle vorhanden
    /// - `Ok(None)` wenn keine Border für diese Zelle
    pub fn get_border_for_cell(&self, row: u32, col: u16) -> Result<Option<Format>, BorderError> {
        // KRITISCH: Merged Cells müssen gesetzt sein!
        let merged_cells = self
            .merged_cells
            .as_ref()
            .ok_or(BorderError::MergedCellsNotSet)?;

        // FEHLER: Borders dürfen NICHT auf non-topleft merged cells gesetzt werden!
        if merged_cells.is_merged_non_topleft(row, col) {
            return Err(BorderError::BorderOnMergedNonTopLeft { row, col });
        }

        let mut border_fmt = Format::new();
        let mut has_border = false;

        // Wenn diese Zelle top-left einer Merge ist, nutze Merge-Außenkanten
        let (effective_start_row, effective_start_col, effective_end_row, effective_end_col) =
            if let Some((sr, sc, er, ec)) = merged_cells.get_merge_range(row, col) {
                (sr, sc, er, ec)
            } else {
                (row, col, row, col)
            };

        for range in &self.ranges {
            if effective_start_row >= range.start_row
                && effective_end_row <= range.end_row
                && effective_start_col >= range.start_col
                && effective_end_col <= range.end_col
            {
                // Edge detection (auf MERGE-Außenkanten, nicht einzelne Zelle!)
                let is_top = effective_start_row == range.start_row;
                let is_bottom = effective_end_row == range.end_row;
                let is_left = effective_start_col == range.start_col;
                let is_right = effective_end_col == range.end_col;

                if is_top {
                    if let Some(border) = range.spec.top {
                        border_fmt = border_fmt.set_border_top(border);
                        has_border = true;
                    }
                }
                if is_bottom {
                    if let Some(border) = range.spec.bottom {
                        border_fmt = border_fmt.set_border_bottom(border);
                        has_border = true;
                    }
                }
                if is_left {
                    if let Some(border) = range.spec.left {
                        border_fmt = border_fmt.set_border_left(border);
                        has_border = true;
                    }
                }
                if is_right {
                    if let Some(border) = range.spec.right {
                        border_fmt = border_fmt.set_border_right(border);
                        has_border = true;
                    }
                }
            }
        }

        if has_border {
            Ok(Some(border_fmt))
        } else {
            Ok(None)
        }
    }
}

impl Default for BorderManager {
    fn default() -> Self {
        Self::new()
    }
}

/// Fehler beim Border-System
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum BorderError {
    /// Merged Cell Registry wurde nicht gesetzt
    MergedCellsNotSet,
    /// Versuch, Border auf non-topleft merged cell zu setzen
    BorderOnMergedNonTopLeft { row: u32, col: u16 },
}

impl fmt::Display for BorderError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            BorderError::MergedCellsNotSet => {
                write!(
                    f,
                    "Merged cell registry not set - call set_merged_cells() first"
                )
            }
            BorderError::BorderOnMergedNonTopLeft { row, col } => {
                write!(
                    f,
                    "Cannot set border on non-topleft merged cell at row {}, col {}",
                    row, col
                )
            }
        }
    }
}

impl std::error::Error for BorderError {}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_border_spec_all_thin() {
        let spec = BorderSpec::all_thin();
        assert!(spec.top.is_some());
        assert!(spec.bottom.is_some());
        assert!(spec.left.is_some());
        assert!(spec.right.is_some());
    }

    #[test]
    fn test_border_spec_top_bottom() {
        let spec = BorderSpec::top_bottom(FormatBorder::Medium);
        assert!(spec.top.is_some());
        assert!(spec.bottom.is_some());
        assert!(spec.left.is_none());
        assert!(spec.right.is_none());
    }

    #[test]
    fn test_merged_cell_registry() {
        let mut registry = MergedCellRegistry::new();
        registry.register_merge(0, 4, 1, 4); // E1:E2

        assert_eq!(registry.count(), 1);

        // E1 (0, 4) ist top-left → nicht merged_non_topleft
        assert!(!registry.is_merged_non_topleft(0, 4));

        // E2 (1, 4) ist non-topleft → ist merged_non_topleft
        assert!(registry.is_merged_non_topleft(1, 4));

        // E1 sollte Merge-Range zurückgeben
        assert_eq!(registry.get_merge_range(0, 4), Some((0, 4, 1, 4)));

        // E2 sollte None zurückgeben (nicht top-left)
        assert_eq!(registry.get_merge_range(1, 4), None);
    }

    #[test]
    fn test_border_manager_without_merged_cells_error() {
        let borders = BorderManager::new();

        // Sollte Fehler werfen, da merged_cells nicht gesetzt
        let result = borders.get_border_for_cell(0, 0);
        assert!(result.is_err());
        assert_eq!(result.unwrap_err(), BorderError::MergedCellsNotSet);
    }

    #[test]
    fn test_border_manager_with_merged_cells() {
        let mut merged_cells = MergedCellRegistry::new();
        merged_cells.register_merge(0, 4, 1, 4); // E1:E2

        let mut borders = BorderManager::new();
        borders.set_merged_cells(merged_cells);
        borders.add_range(0, 4, 1, 4, BorderSpec::all_medium());

        // E1 (top-left) sollte Border bekommen
        let result = borders.get_border_for_cell(0, 4);
        assert!(result.is_ok());
        assert!(result.unwrap().is_some());

        // E2 (non-topleft) sollte Fehler werfen
        let result = borders.get_border_for_cell(1, 4);
        assert!(result.is_err());
        assert_eq!(
            result.unwrap_err(),
            BorderError::BorderOnMergedNonTopLeft { row: 1, col: 4 }
        );
    }

    #[test]
    fn test_border_manager_normal_cell() {
        let merged_cells = MergedCellRegistry::new(); // Leer

        let mut borders = BorderManager::new();
        borders.set_merged_cells(merged_cells);
        borders.add_range(0, 0, 5, 5, BorderSpec::all_thin());

        // Zelle (0, 0) ist top-left corner
        let result = borders.get_border_for_cell(0, 0);
        assert!(result.is_ok());
        assert!(result.unwrap().is_some());

        // Zelle (2, 2) ist in der Mitte (keine Border außer wenn Bereich abdeckt)
        let result = borders.get_border_for_cell(2, 2);
        assert!(result.is_ok());
        assert!(result.unwrap().is_none()); // Nicht am Rand
    }
}
