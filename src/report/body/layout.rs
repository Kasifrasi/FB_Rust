//! Body-Layout Berechnung
//!
//! Berechnet die Zeilen-Positionen basierend auf BodyConfig.
//!
//! ## Layout-Modi
//!
//! Jede Kategorie kann in einem von zwei Modi sein:
//!
//! ### Header-Eingabe-Modus (0 Positionen)
//! ```text
//! | B     | C (Label)      | D        | E        | F        | G     | H        |
//! |-------|----------------|----------|----------|----------|-------|----------|
//! | 6.    | VLOOKUP Label  | [Input]  | [Input]  | [Input]  | Ratio | [Input]  |
//! ```
//! - Nur eine Zeile
//! - Label in C (VLOOKUP)
//! - Eingabe direkt in D, E, F, H
//! - API: `Position { category: 6, position: 0, field }` (position=0!)
//!
//! ### Positions-Modus (1+ Positionen)
//! ```text
//! | B     | C              | D        | E        | F        | G     | H        |
//! |-------|----------------|----------|----------|----------|-------|----------|
//! | 1.    | VLOOKUP Label  |          |          |          |       |          | <- Header
//! | 1.1   | [Input]        | [Input]  | [Input]  | [Input]  | Ratio | [Input]  | <- Position 1
//! | 1.2   | [Input]        | [Input]  | [Input]  | [Input]  | Ratio | [Input]  | <- Position 2
//! | ...   | ...            | ...      | ...      | ...      | ...   | ...      |
//! |       | VLOOKUP Sum    | SUM      | SUM      | SUM      | Ratio |        | <- Footer
//! ```
//! - Header-Zeile (Kategorie-Label, keine Eingabe)
//! - N Positions-Zeilen (Eingabe)
//! - Footer-Zeile (Summen)
//! - API: `Position { category: 1, position: 1..N, field }` (position 1-basiert!)

use super::config::{BodyConfig, ALL_CATEGORIES, BODY_START_ROW};
use crate::report::api::PositionField;
use crate::report::core::CellAddr;

/// Kategorie-Metadaten (VLOOKUP-Indizes)
#[derive(Debug, Clone, Copy)]
pub struct CategoryMeta {
    /// Kategorie-Nummer (1-8)
    pub num: u8,
    /// VLOOKUP-Index für Header/Kategorie-Label
    pub label_index: usize,
    /// VLOOKUP-Index für Footer/Sum-Label (nur bei Positions-Modus)
    pub sum_label_index: usize,
}

/// Alle Kategorie-Metadaten
/// VLOOKUP-Indizes aus Sprachversionen-Sheet:
/// - Kategorien 1-5: Label + Sum-Label Paare (29/30, 31/32, 33/34, 35/36, 37/38)
/// - Kategorie 6 (Evaluierung): Label 39, Sum-Label 64
/// - Kategorie 7 (Audit): Label 40, Sum-Label 65
/// - Kategorie 8 (Reserve): Label 41, Sum-Label 66
/// - Gesamt: Index 42
pub const CATEGORY_META: [CategoryMeta; 8] = [
    CategoryMeta {
        num: 1,
        label_index: 29,
        sum_label_index: 30,
    },
    CategoryMeta {
        num: 2,
        label_index: 31,
        sum_label_index: 32,
    },
    CategoryMeta {
        num: 3,
        label_index: 33,
        sum_label_index: 34,
    },
    CategoryMeta {
        num: 4,
        label_index: 35,
        sum_label_index: 36,
    },
    CategoryMeta {
        num: 5,
        label_index: 37,
        sum_label_index: 38,
    },
    CategoryMeta {
        num: 6,
        label_index: 39,
        sum_label_index: 64, // Zwischensumme Evaluierung
    },
    CategoryMeta {
        num: 7,
        label_index: 40,
        sum_label_index: 65, // Zwischensumme Audit
    },
    CategoryMeta {
        num: 8,
        label_index: 41,
        sum_label_index: 66, // Zwischensumme Reserve
    },
];

/// VLOOKUP-Index für "Gesamt" Label
pub const TOTAL_LABEL_INDEX: usize = 42;

/// Berechnetes Body-Layout
#[derive(Debug, Clone)]
pub struct BodyLayout {
    /// Layout für jede Kategorie
    pub categories: Vec<CategoryLayout>,
    /// Zeile der Gesamt-Summe
    pub total_row: u32,
    /// Letzte beschriebene Zeile
    pub last_row: u32,
}

/// Layout einer einzelnen Kategorie
#[derive(Debug, Clone)]
pub struct CategoryLayout {
    /// Kategorie-Metadaten
    pub meta: CategoryMeta,
    /// Modus der Kategorie
    pub mode: CategoryMode,
}

/// Modus einer Kategorie
#[derive(Debug, Clone)]
pub enum CategoryMode {
    /// Header-Eingabe-Modus (0 Positionen)
    /// Die Header-Zeile ist gleichzeitig die Eingabezeile
    HeaderInput {
        /// Die Zeile (Header = Eingabe)
        row: u32,
    },
    /// Positions-Modus (1+ Positionen)
    /// Header, N Positionen, Footer
    WithPositions {
        /// Header-Zeile
        header_row: u32,
        /// Positions-Bereich
        positions: PositionRange,
        /// Footer-Zeile
        footer_row: u32,
    },
}

/// Bereich der Positions-Zeilen
#[derive(Debug, Clone, Copy)]
pub struct PositionRange {
    /// Erste Positions-Zeile (0-basiert)
    pub start_row: u32,
    /// Letzte Positions-Zeile (0-basiert)
    pub end_row: u32,
    /// Anzahl Positionen
    pub count: u16,
}

impl BodyLayout {
    /// Berechnet das Layout aus der Konfiguration
    pub fn compute(config: &BodyConfig) -> Self {
        let mut current_row = BODY_START_ROW;
        let mut categories = Vec::with_capacity(8);

        for &cat_num in &ALL_CATEGORIES {
            let meta = CATEGORY_META[(cat_num - 1) as usize];
            let num_positions = config.position_count(cat_num);

            if num_positions == 0 {
                // Header-Eingabe-Modus: nur eine Zeile
                categories.push(CategoryLayout {
                    meta,
                    mode: CategoryMode::HeaderInput { row: current_row },
                });
                current_row += 1;
            } else {
                // Positions-Modus: Header + N Positionen + Footer
                let header_row = current_row;
                current_row += 1;

                let positions_start = current_row;
                let positions_end = current_row + num_positions as u32 - 1;
                current_row += num_positions as u32;

                let footer_row = current_row;
                current_row += 1;

                categories.push(CategoryLayout {
                    meta,
                    mode: CategoryMode::WithPositions {
                        header_row,
                        positions: PositionRange {
                            start_row: positions_start,
                            end_row: positions_end,
                            count: num_positions,
                        },
                        footer_row,
                    },
                });
            }
        }

        let total_row = current_row;

        Self {
            categories,
            total_row,
            last_row: total_row,
        }
    }

    /// Gibt das Layout für eine Kategorie zurück
    pub fn category(&self, num: u8) -> Option<&CategoryLayout> {
        self.categories.iter().find(|c| c.meta.num == num)
    }

    /// Gibt alle Footer-Zeilen zurück (für Gesamt-Summe)
    pub fn footer_rows(&self) -> Vec<u32> {
        self.categories
            .iter()
            .filter_map(|c| c.footer_row())
            .collect()
    }

    /// Gibt alle Header-Input-Zeilen zurück (für Gesamt-Summe)
    pub fn header_input_rows(&self) -> Vec<u32> {
        self.categories
            .iter()
            .filter_map(|c| c.header_input_row())
            .collect()
    }

    /// Gibt alle Zeilen zurück die eine Ratio-Formel brauchen
    pub fn ratio_rows(&self) -> Vec<u32> {
        let mut rows = Vec::new();

        for cat in &self.categories {
            match &cat.mode {
                CategoryMode::HeaderInput { row } => {
                    rows.push(*row);
                }
                CategoryMode::WithPositions {
                    positions,
                    footer_row,
                    ..
                } => {
                    // Positions-Zeilen
                    for row in positions.start_row..=positions.end_row {
                        rows.push(row);
                    }
                    // Footer-Zeile
                    rows.push(*footer_row);
                }
            }
        }

        // Gesamt-Zeile
        rows.push(self.total_row);

        rows
    }

    /// Gesamtanzahl der Zeilen im Body
    pub fn row_count(&self) -> u32 {
        self.last_row - BODY_START_ROW + 1
    }

    /// Berechnet die Zelladresse für ein Positions-Feld
    ///
    /// # Arguments
    /// * `category` - Kategorie-Nummer (1-8)
    /// * `position` - Position:
    ///   - `0`: Header-Eingabe (nur bei Header-Input-Modus!)
    ///   - `1..N`: Positions-Zeile (nur bei Positions-Modus!)
    /// * `field` - Welches Feld der Position
    ///
    /// # Returns
    /// `Some(CellAddr)` wenn gültig, `None` wenn:
    /// - Kategorie nicht existiert
    /// - position=0 aber Kategorie hat Positionen
    /// - position>0 aber Kategorie ist Header-Input
    /// - Position außerhalb des Bereichs
    ///
    /// # Hinweis zu PositionField::Description bei Header-Input
    /// Bei Header-Input-Modus (position=0) gibt Description `None` zurück,
    /// da Spalte C das VLOOKUP-Label enthält und nicht eingegeben werden kann.
    pub fn position_addr(
        &self,
        category: u8,
        position: u16,
        field: PositionField,
    ) -> Option<CellAddr> {
        let cat = self.categories.iter().find(|c| c.meta.num == category)?;

        match &cat.mode {
            CategoryMode::HeaderInput { row } => {
                // Header-Input: nur position=0 erlaubt
                if position != 0 {
                    return None;
                }
                // Description nicht erlaubt (C ist VLOOKUP-Label)
                if matches!(field, PositionField::Description) {
                    return None;
                }
                Some(CellAddr::new(*row, field.col()))
            }
            CategoryMode::WithPositions { positions, .. } => {
                // Positions-Modus: position muss 1..N sein
                if position < 1 || position > positions.count {
                    return None;
                }
                let row = positions.start_row + (position - 1) as u32;
                Some(CellAddr::new(row, field.col()))
            }
        }
    }
}

impl CategoryLayout {
    /// Gibt die Header-Zeile zurück (falls im Positions-Modus)
    pub fn header_row(&self) -> Option<u32> {
        match &self.mode {
            CategoryMode::WithPositions { header_row, .. } => Some(*header_row),
            CategoryMode::HeaderInput { .. } => None,
        }
    }

    /// Gibt die Footer-Zeile zurück (falls im Positions-Modus)
    pub fn footer_row(&self) -> Option<u32> {
        match &self.mode {
            CategoryMode::WithPositions { footer_row, .. } => Some(*footer_row),
            CategoryMode::HeaderInput { .. } => None,
        }
    }

    /// Gibt die Header-Input-Zeile zurück (falls im Header-Input-Modus)
    pub fn header_input_row(&self) -> Option<u32> {
        match &self.mode {
            CategoryMode::HeaderInput { row } => Some(*row),
            CategoryMode::WithPositions { .. } => None,
        }
    }

    /// Gibt den Positions-Bereich zurück (falls im Positions-Modus)
    pub fn positions(&self) -> Option<&PositionRange> {
        match &self.mode {
            CategoryMode::WithPositions { positions, .. } => Some(positions),
            CategoryMode::HeaderInput { .. } => None,
        }
    }

    /// Ist dies eine Header-Input-Kategorie?
    pub fn is_header_input(&self) -> bool {
        matches!(self.mode, CategoryMode::HeaderInput { .. })
    }

    /// Hat diese Kategorie Positionen?
    pub fn has_positions(&self) -> bool {
        matches!(self.mode, CategoryMode::WithPositions { .. })
    }

    /// Gibt die "Summen-Zeile" zurück (Footer bei Positions-Modus, Header bei Header-Input)
    /// Diese Zeile wird in der Gesamt-Summe referenziert.
    pub fn sum_row(&self) -> u32 {
        match &self.mode {
            CategoryMode::HeaderInput { row } => *row,
            CategoryMode::WithPositions { footer_row, .. } => *footer_row,
        }
    }
}

/// Footer-Layout mit berechneten Zeilenpositionen
#[derive(Debug, Clone)]
pub struct FooterLayout {
    /// Startzeile des Footers (0-indexed)
    pub start_row: u32,
    /// Zeile für Saldo-Formel (E-Spalte)
    pub saldo_row: u32,
    /// Zeilen für Bank/Kasse/Sonstiges Input (E-Spalte)
    pub input_rows: [u32; 3],
    /// Letzte Zeile des Footers
    pub end_row: u32,
}

impl FooterLayout {
    /// Berechnet das Footer-Layout basierend auf der Total-Zeile des Body
    ///
    /// Der Footer beginnt 3 Zeilen nach dem Total.
    pub fn compute(total_row: u32) -> Self {
        let start_row = total_row + 3; // 3 Zeilen Abstand

        Self {
            start_row,
            saldo_row: start_row + 4, // Zeile 4: Saldo-Differenz
            input_rows: [
                start_row + 7, // Zeile 7: Bank
                start_row + 8, // Zeile 8: Kasse
                start_row + 9, // Zeile 9: Sonstiges
            ],
            end_row: start_row + 20, // 21 Zeilen (0-20)
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_layout_with_defaults() {
        let config = BodyConfig::default();
        let layout = BodyLayout::compute(&config);

        // Kategorie 1: Positions-Modus (20 Positionen)
        let cat1 = layout.category(1).unwrap();
        assert!(cat1.has_positions());
        assert!(!cat1.is_header_input());

        if let CategoryMode::WithPositions {
            header_row,
            positions,
            footer_row,
        } = &cat1.mode
        {
            assert_eq!(*header_row, 26);
            assert_eq!(positions.start_row, 27);
            assert_eq!(positions.end_row, 46);
            assert_eq!(positions.count, 20);
            assert_eq!(*footer_row, 47);
        }

        // Kategorie 6: Header-Input-Modus (0 Positionen)
        let cat6 = layout.category(6).unwrap();
        assert!(cat6.is_header_input());
        assert!(!cat6.has_positions());
    }

    #[test]
    fn test_layout_with_custom_config() {
        let config = BodyConfig::new()
            .with_positions(1, 5)
            .with_positions(2, 10)
            .with_positions(6, 3); // Kategorie 6 jetzt mit Positionen!

        let layout = BodyLayout::compute(&config);

        // Kategorie 1: 5 Positionen
        let cat1 = layout.category(1).unwrap();
        assert_eq!(cat1.positions().unwrap().count, 5);

        // Kategorie 6: jetzt mit 3 Positionen!
        let cat6 = layout.category(6).unwrap();
        assert!(cat6.has_positions());
        assert_eq!(cat6.positions().unwrap().count, 3);
    }

    #[test]
    fn test_header_input_mode() {
        let config = BodyConfig::new()
            .with_positions(1, 0) // Kategorie 1 im Header-Input-Modus!
            .with_positions(2, 5);

        let layout = BodyLayout::compute(&config);

        // Kategorie 1: Header-Input
        let cat1 = layout.category(1).unwrap();
        assert!(cat1.is_header_input());
        assert_eq!(cat1.header_input_row(), Some(26));

        // Kategorie 2: Positions-Modus, startet direkt nach Kategorie 1
        let cat2 = layout.category(2).unwrap();
        assert!(cat2.has_positions());
        assert_eq!(cat2.header_row(), Some(27)); // Direkt nach Kategorie 1
    }

    #[test]
    fn test_position_addr_header_input() {
        use PositionField::*;

        let config = BodyConfig::new().with_positions(1, 0); // Header-Input
        let layout = BodyLayout::compute(&config);

        // position=0 bei Header-Input gibt Adresse zurück
        let addr = layout.position_addr(1, 0, Approved).unwrap();
        assert_eq!(addr.row, 26);
        assert_eq!(addr.col, 3); // D

        // Description nicht erlaubt bei Header-Input (C ist VLOOKUP)
        assert!(layout.position_addr(1, 0, Description).is_none());

        // position>0 nicht erlaubt bei Header-Input
        assert!(layout.position_addr(1, 1, Approved).is_none());
    }

    #[test]
    fn test_position_addr_with_positions() {
        use PositionField::*;

        let config = BodyConfig::new().with_positions(1, 5);
        let layout = BodyLayout::compute(&config);

        // position=0 nicht erlaubt bei Positions-Modus
        assert!(layout.position_addr(1, 0, Approved).is_none());

        // position=1..5 erlaubt
        let addr = layout.position_addr(1, 1, Description).unwrap();
        assert_eq!(addr.row, 27);
        assert_eq!(addr.col, 2); // C

        let addr = layout.position_addr(1, 5, IncomeTotal).unwrap();
        assert_eq!(addr.row, 31);
        assert_eq!(addr.col, 5); // F

        // position=6 nicht erlaubt (nur 5 Positionen)
        assert!(layout.position_addr(1, 6, Approved).is_none());
    }

    #[test]
    fn test_sum_rows() {
        let config = BodyConfig::new()
            .with_positions(1, 2)
            .with_positions(2, 0) // Header-Input
            .with_positions(3, 3);

        let layout = BodyLayout::compute(&config);

        // Kategorie 1: Footer-Zeile ist Summen-Zeile
        let cat1 = layout.category(1).unwrap();
        assert_eq!(cat1.sum_row(), cat1.footer_row().unwrap());

        // Kategorie 2: Header-Input-Zeile ist Summen-Zeile
        let cat2 = layout.category(2).unwrap();
        assert_eq!(cat2.sum_row(), cat2.header_input_row().unwrap());
    }

    #[test]
    fn test_ratio_rows() {
        let config = BodyConfig::new()
            .with_positions(1, 2)
            .with_positions(2, 0) // Header-Input
            .with_positions(3, 0) // Header-Input
            .with_positions(4, 1)
            .with_positions(5, 0)
            .with_positions(6, 0)
            .with_positions(7, 0)
            .with_positions(8, 0);

        let layout = BodyLayout::compute(&config);
        let ratio_rows = layout.ratio_rows();

        // Kat 1: 2 Positionen + 1 Footer = 3
        // Kat 2: 1 Header-Input = 1
        // Kat 3: 1 Header-Input = 1
        // Kat 4: 1 Position + 1 Footer = 2
        // Kat 5-8: je 1 Header-Input = 4
        // + 1 Total = 1
        // Gesamt: 3 + 1 + 1 + 2 + 4 + 1 = 12
        assert_eq!(ratio_rows.len(), 12);
    }
}
