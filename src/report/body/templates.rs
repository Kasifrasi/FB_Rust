//! Zeilenbasierte Templates für dynamischen Body-Bereich
//!
//! Dieses Modul implementiert das Baukasten-Prinzip:
//! - Templates beschreiben komplette Zeilen (Format + API-Felder)
//! - Format und API bleiben GETRENNT (modular)
//! - Template ist nur die KOORDINATION ("Spalte C → FillKey::Input + PositionField::Description")
//!
//! # Beispiel
//!
//! ```ignore
//! let template = RowTemplate::position_row();
//!
//! // Template koordiniert Format (fills.rs) + API (registry.rs)
//! template.apply_to_row(
//!     ws,
//!     registry,
//!     fills,
//!     row,
//!     category,
//!     position,
//! )?;
//! ```

use crate::report::api::{ApiKey, PositionField};
use crate::report::core::{CellAddr, CellRegistry};
use crate::report::format::{BorderManager, FillColors};
use rust_xlsxwriter::{Format, FormatAlign, Worksheet, XlsxError};

/// Fill-Keys (verweisen auf FillColors)
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FillKey {
    /// Input-Felder (gelb)
    Input,
    /// Berechnete Werte (grau hell)
    Value,
    /// Summary-Zeilen (grau dunkel)
    Summary,
    /// Highlight/Special (orange)
    Highlight,
}

impl FillKey {
    /// Gibt die entsprechende Farbe aus FillColors zurück
    pub fn color(&self, fills: &FillColors) -> rust_xlsxwriter::Color {
        match self {
            FillKey::Input => fills.input,
            FillKey::Value => fills.value,
            FillKey::Summary => fills.summary,
            FillKey::Highlight => fills.highlight,
        }
    }
}

/// Spezifikation für eine Zelle im Template
#[derive(Debug, Clone)]
pub struct CellSpec {
    /// Spalte (B=1, C=2, D=3, ...)
    pub col: u16,
    /// Welche Fill-Variable? (optional)
    pub fill_key: Option<FillKey>,
    /// Welches API-Feld? (optional)
    pub api_field: Option<PositionField>,

    // Format-Properties (explizit pro Zelle)
    pub font: &'static str,
    pub size: f64,
    pub bold: bool,
    pub align: FormatAlign,
    pub num_format: &'static str,
    pub locked: bool,
}

impl CellSpec {
    /// Erstellt ein Format-Objekt aus dieser Spec + Fill-Farbe
    pub fn to_format(&self, fills: &FillColors) -> Format {
        let mut fmt = Format::new()
            .set_font_name(self.font)
            .set_font_size(self.size)
            .set_align(self.align)
            .set_num_format(self.num_format);

        if self.bold {
            fmt = fmt.set_bold();
        }

        if !self.locked {
            fmt = fmt.set_unlocked();
        }

        if let Some(fill_key) = self.fill_key {
            fmt = fmt.set_background_color(fill_key.color(fills));
        }

        fmt
    }
}

/// Zeilenbasiertes Template
///
/// WICHTIG: Format und API bleiben GETRENNT (modular)!
/// Template ist nur die Koordination: "Welche Spalte bekommt was?"
#[derive(Debug, Clone)]
pub struct RowTemplate {
    pub row_type: RowType,
    pub specs: Vec<CellSpec>,
}

/// Row-Typ für Templates
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RowType {
    /// Position-Zeile (Input-Zeile im Body)
    Position,
    /// Category Header (wenn Positionen existieren)
    CategoryHeader,
    /// Footer-Zeile (SUMPRODUCT)
    Footer,
    /// Total-Zeile (Grand total)
    Total,
}

impl RowTemplate {
    /// Template für Position-Zeile im Body
    ///
    /// Spalten:
    /// - B: Position-Nummer (z.B. "1.1", "1.2") - statisch, locked
    /// - C: Beschreibung - Input, gelb, unlocked, API
    /// - D: Bewilligt - Input, gelb, unlocked, API
    /// - E: Einnahmen Bericht - Input, gelb, unlocked, API
    /// - F: Einnahmen Total - Input, gelb, unlocked, API
    /// - G: Ratio - Formel, locked
    /// - H: Bemerkung - Input, gelb, unlocked, API
    pub fn position_row() -> Self {
        Self {
            row_type: RowType::Position,
            specs: vec![
                // B: Position-Nummer (statisch, "1.1", "1.2", etc.)
                CellSpec {
                    col: 1,
                    fill_key: None,
                    api_field: None,
                    font: "Arial",
                    size: 10.0,
                    bold: false,
                    align: FormatAlign::Right,
                    num_format: "@",
                    locked: true,
                },
                // C: Beschreibung (Input, text)
                CellSpec {
                    col: 2,
                    fill_key: Some(FillKey::Input),
                    api_field: Some(PositionField::Description),
                    font: "Arial",
                    size: 10.0,
                    bold: false,
                    align: FormatAlign::Left,
                    num_format: "@",
                    locked: false,
                },
                // D: Bewilligt (Input, number)
                CellSpec {
                    col: 3,
                    fill_key: Some(FillKey::Input),
                    api_field: Some(PositionField::Approved),
                    font: "Arial",
                    size: 10.0,
                    bold: false,
                    align: FormatAlign::Right,
                    num_format: "#,##0.00",
                    locked: false,
                },
                // E: Einnahmen Bericht (Input, number)
                CellSpec {
                    col: 4,
                    fill_key: Some(FillKey::Input),
                    api_field: Some(PositionField::IncomeReport),
                    font: "Arial",
                    size: 10.0,
                    bold: false,
                    align: FormatAlign::Right,
                    num_format: "#,##0.00",
                    locked: false,
                },
                // F: Einnahmen Total (Input, number)
                CellSpec {
                    col: 5,
                    fill_key: Some(FillKey::Input),
                    api_field: Some(PositionField::IncomeTotal),
                    font: "Arial",
                    size: 10.0,
                    bold: false,
                    align: FormatAlign::Right,
                    num_format: "#,##0.00",
                    locked: false,
                },
                // G: Ratio (Formula, percentage)
                CellSpec {
                    col: 6,
                    fill_key: None,
                    api_field: None,
                    font: "Arial",
                    size: 10.0,
                    bold: false,
                    align: FormatAlign::Center,
                    num_format: "0.00%",
                    locked: true,
                },
                // H: Bemerkung (Input, text)
                CellSpec {
                    col: 7,
                    fill_key: Some(FillKey::Input),
                    api_field: Some(PositionField::Remark),
                    font: "Arial",
                    size: 10.0,
                    bold: false,
                    align: FormatAlign::Left,
                    num_format: "@",
                    locked: false,
                },
            ],
        }
    }

    /// Template für Category Header-Zeile
    ///
    /// Header-Zeile wenn Kategorie Positionen hat
    pub fn category_header() -> Self {
        Self {
            row_type: RowType::CategoryHeader,
            specs: vec![
                // B: Kategorie-Nummer (z.B. "1.")
                CellSpec {
                    col: 1,
                    fill_key: None,
                    api_field: None,
                    font: "Arial",
                    size: 10.0,
                    bold: true,
                    align: FormatAlign::Right,
                    num_format: "@",
                    locked: true,
                },
                // C: Kategorie-Name (VLOOKUP)
                CellSpec {
                    col: 2,
                    fill_key: None,
                    api_field: None,
                    font: "Arial",
                    size: 10.0,
                    bold: true,
                    align: FormatAlign::Left,
                    num_format: "@",
                    locked: true,
                },
                // D-H: Leer
                CellSpec {
                    col: 3,
                    fill_key: None,
                    api_field: None,
                    font: "Arial",
                    size: 10.0,
                    bold: false,
                    align: FormatAlign::Right,
                    num_format: "General",
                    locked: true,
                },
            ],
        }
    }

    /// Template für Footer-Zeile (Kategorie-Summe)
    pub fn category_footer() -> Self {
        Self {
            row_type: RowType::Footer,
            specs: vec![
                // B: Leer
                CellSpec {
                    col: 1,
                    fill_key: None,
                    api_field: None,
                    font: "Arial",
                    size: 10.0,
                    bold: false,
                    align: FormatAlign::Right,
                    num_format: "General",
                    locked: true,
                },
                // C: "Zwischensumme" (VLOOKUP)
                CellSpec {
                    col: 2,
                    fill_key: None,
                    api_field: None,
                    font: "Arial",
                    size: 10.0,
                    bold: true,
                    align: FormatAlign::Left,
                    num_format: "@",
                    locked: true,
                },
                // D-F: SUMPRODUCT formulas (grau)
                CellSpec {
                    col: 3,
                    fill_key: Some(FillKey::Value),
                    api_field: None,
                    font: "Arial",
                    size: 10.0,
                    bold: true,
                    align: FormatAlign::Right,
                    num_format: "#,##0.00",
                    locked: true,
                },
                // G: Ratio
                CellSpec {
                    col: 6,
                    fill_key: None,
                    api_field: None,
                    font: "Arial",
                    size: 10.0,
                    bold: true,
                    align: FormatAlign::Center,
                    num_format: "0.00%",
                    locked: true,
                },
            ],
        }
    }

    /// Wendet Template auf eine Zeile an
    ///
    /// MODULAR: Format und API bleiben getrennt!
    /// - Format: fills.rs (FillColors) + Template (explizite Properties) + borders.rs (BorderManager)
    /// - API: registry.rs (register_api_at)
    /// - Template: Nur Koordination ("Spalte C bekommt Fill Input + API Description")
    pub fn apply_to_row<E>(
        &self,
        ws: &mut Worksheet,
        registry: &mut CellRegistry<E>,
        fills: &FillColors,
        borders: Option<&BorderManager>,
        row: u32,
        category: u8,
        position: u16,
    ) -> Result<(), XlsxError> {
        for spec in &self.specs {
            // 1. Baue Format aus Template-Spec + Fill-Variable
            let mut fmt = spec.to_format(fills);

            // 2. Merge mit Borders falls vorhanden (Baukastenprinzip!)
            if let Some(border_mgr) = borders {
                if let Ok(Some(border_fmt)) = border_mgr.get_border_for_cell(row, spec.col) {
                    fmt = fmt.merge(&border_fmt);
                }
            }

            // 3. Schreibe Zelle mit Format
            ws.write_blank(row, spec.col, &fmt)?;

            // 4. Registriere API-Feld (modular, getrennt!)
            if let Some(field) = spec.api_field {
                let api_key = ApiKey::Position {
                    category,
                    position,
                    field,
                };

                let addr = CellAddr::new(row, spec.col);
                registry.register_api_at(api_key, addr).ok(); // Ignoriere Fehler wenn bereits registriert
            }
        }

        Ok(())
    }

    /// Wendet Template auf eine Zeile an (ohne API-Registrierung)
    ///
    /// Für Zeilen die keine API-Felder haben (Header, Footer, etc.)
    pub fn apply_to_row_no_api(
        &self,
        ws: &mut Worksheet,
        fills: &FillColors,
        borders: Option<&BorderManager>,
        row: u32,
    ) -> Result<(), XlsxError> {
        for spec in &self.specs {
            let mut fmt = spec.to_format(fills);

            // Merge mit Borders falls vorhanden (Baukastenprinzip!)
            if let Some(border_mgr) = borders {
                if let Ok(Some(border_fmt)) = border_mgr.get_border_for_cell(row, spec.col) {
                    fmt = fmt.merge(&border_fmt);
                }
            }

            ws.write_blank(row, spec.col, &fmt)?;
        }

        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_fill_key_colors() {
        let fills = FillColors::new();

        assert_eq!(FillKey::Input.color(&fills), fills.input);
        assert_eq!(FillKey::Value.color(&fills), fills.value);
        assert_eq!(FillKey::Summary.color(&fills), fills.summary);
        assert_eq!(FillKey::Highlight.color(&fills), fills.highlight);
    }

    #[test]
    fn test_position_row_template() {
        let template = RowTemplate::position_row();

        assert_eq!(template.row_type, RowType::Position);
        assert_eq!(template.specs.len(), 7); // B-H

        // B: Position-Nummer (locked, no fill, no API)
        assert_eq!(template.specs[0].col, 1);
        assert!(template.specs[0].locked);
        assert!(template.specs[0].fill_key.is_none());
        assert!(template.specs[0].api_field.is_none());

        // C: Beschreibung (unlocked, Input fill, Description API)
        assert_eq!(template.specs[1].col, 2);
        assert!(!template.specs[1].locked);
        assert_eq!(template.specs[1].fill_key, Some(FillKey::Input));
        assert_eq!(
            template.specs[1].api_field,
            Some(PositionField::Description)
        );

        // D: Bewilligt (unlocked, Input fill, Approved API)
        assert_eq!(template.specs[2].col, 3);
        assert!(!template.specs[2].locked);
        assert_eq!(template.specs[2].fill_key, Some(FillKey::Input));
        assert_eq!(template.specs[2].api_field, Some(PositionField::Approved));
    }

    #[test]
    fn test_cell_spec_to_format() {
        let fills = FillColors::new();

        let spec = CellSpec {
            col: 2,
            fill_key: Some(FillKey::Input),
            api_field: Some(PositionField::Description),
            font: "Arial",
            size: 10.0,
            bold: false,
            align: FormatAlign::Left,
            num_format: "@",
            locked: false,
        };

        let fmt = spec.to_format(&fills);
        // Format wurde erstellt (weitere Assertions schwer da Format private ist)
    }

    #[test]
    fn test_category_header_template() {
        let template = RowTemplate::category_header();

        assert_eq!(template.row_type, RowType::CategoryHeader);
        assert_eq!(template.specs.len(), 3); // B, C, D (simplified)

        // B: Bold, locked
        assert!(template.specs[0].bold);
        assert!(template.specs[0].locked);

        // C: Bold, locked (Kategorie-Name)
        assert!(template.specs[1].bold);
        assert!(template.specs[1].locked);
    }

    #[test]
    fn test_category_footer_template() {
        let template = RowTemplate::category_footer();

        assert_eq!(template.row_type, RowType::Footer);

        // Footer hat Summary-Fill für berechnete Werte
        assert_eq!(template.specs[2].fill_key, Some(FillKey::Value));
    }
}
