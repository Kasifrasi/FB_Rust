//! Body-Layout Berechnung
//!
//! Berechnet die Zeilen-Positionen basierend auf BodyConfig.

use super::config::{BodyConfig, BODY_START_ROW, MULTI_ROW_CATEGORIES};
use crate::v2::report::api::{PositionField, SingleRowField};
use crate::v2::report::registry::CellAddr;

/// Kategorie-Metadaten (VLOOKUP-Indizes)
#[derive(Debug, Clone, Copy)]
pub struct CategoryMeta {
    /// Kategorie-Nummer (1-8)
    pub num: u8,
    /// VLOOKUP-Index für Header-Label
    pub label_index: usize,
    /// VLOOKUP-Index für Footer/Sum-Label (0 = Single-Row)
    pub sum_label_index: usize,
}

/// Alle Kategorie-Metadaten
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
        sum_label_index: 0,
    },
    CategoryMeta {
        num: 7,
        label_index: 40,
        sum_label_index: 0,
    },
    CategoryMeta {
        num: 8,
        label_index: 41,
        sum_label_index: 0,
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
    /// Header-Zeile (nur Multi-Row)
    pub header_row: Option<u32>,
    /// Positions-Bereich (nur Multi-Row)
    pub positions: Option<PositionRange>,
    /// Footer-Zeile (nur Multi-Row)
    pub footer_row: Option<u32>,
    /// Single-Row Zeile (nur Single-Row Kategorien)
    pub single_row: Option<u32>,
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

        for meta in CATEGORY_META {
            let is_multi = MULTI_ROW_CATEGORIES.contains(&meta.num);

            if is_multi {
                // Multi-Row: Header + Positionen + Footer
                let header_row = current_row;
                current_row += 1;

                let num_positions = config.position_count(meta.num);
                let positions_start = current_row;
                let positions_end = current_row + num_positions as u32 - 1;
                current_row += num_positions as u32;

                let footer_row = current_row;
                current_row += 1;

                categories.push(CategoryLayout {
                    meta,
                    header_row: Some(header_row),
                    positions: Some(PositionRange {
                        start_row: positions_start,
                        end_row: positions_end,
                        count: num_positions,
                    }),
                    footer_row: Some(footer_row),
                    single_row: None,
                });
            } else {
                // Single-Row: Nur eine Zeile
                let single_row = current_row;
                current_row += 1;

                categories.push(CategoryLayout {
                    meta,
                    header_row: None,
                    positions: None,
                    footer_row: None,
                    single_row: Some(single_row),
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
            .filter_map(|c| c.footer_row)
            .collect()
    }

    /// Gibt alle Single-Row-Zeilen zurück (für Gesamt-Summe)
    pub fn single_rows(&self) -> Vec<u32> {
        self.categories
            .iter()
            .filter_map(|c| c.single_row)
            .collect()
    }

    /// Gibt alle Zeilen zurück die eine Ratio-Formel brauchen
    pub fn ratio_rows(&self) -> Vec<u32> {
        let mut rows = Vec::new();

        for cat in &self.categories {
            // Positions-Zeilen
            if let Some(positions) = &cat.positions {
                for row in positions.start_row..=positions.end_row {
                    rows.push(row);
                }
            }
            // Footer-Zeilen
            if let Some(footer_row) = cat.footer_row {
                rows.push(footer_row);
            }
            // Single-Row-Zeilen
            if let Some(single_row) = cat.single_row {
                rows.push(single_row);
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
    /// * `category` - Kategorie-Nummer (1-5, Multi-Row Kategorien)
    /// * `position` - Position innerhalb der Kategorie (1-N, 1-basiert!)
    /// * `field` - Welches Feld der Position
    ///
    /// # Returns
    /// `Some(CellAddr)` wenn gültig, `None` wenn:
    /// - Kategorie nicht existiert
    /// - Kategorie keine Positionen hat (Single-Row)
    /// - Position außerhalb des Bereichs
    pub fn position_addr(
        &self,
        category: u8,
        position: u16,
        field: PositionField,
    ) -> Option<CellAddr> {
        let cat = self.categories.iter().find(|c| c.meta.num == category)?;
        let positions = cat.positions.as_ref()?;

        if position < 1 || position > positions.count {
            return None;
        }

        let row = positions.start_row + (position - 1) as u32;
        Some(CellAddr::new(row, field.col()))
    }

    /// Berechnet die Zelladresse für ein Single-Row-Feld
    ///
    /// # Arguments
    /// * `category` - Kategorie-Nummer (6, 7 oder 8)
    /// * `field` - Welches Feld
    ///
    /// # Returns
    /// `Some(CellAddr)` wenn gültig, `None` wenn:
    /// - Kategorie nicht existiert
    /// - Kategorie keine Single-Row ist
    pub fn single_row_addr(&self, category: u8, field: SingleRowField) -> Option<CellAddr> {
        let cat = self.categories.iter().find(|c| c.meta.num == category)?;
        let row = cat.single_row?;
        Some(CellAddr::new(row, field.col()))
    }
}

impl CategoryLayout {
    /// Ist dies eine Multi-Row Kategorie?
    pub fn is_multi_row(&self) -> bool {
        self.positions.is_some()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_layout_with_defaults() {
        let config = BodyConfig::default();
        let layout = BodyLayout::compute(&config);

        // Kategorie 1: Row 26 (Header), 27-46 (20 Positions), 47 (Footer)
        let cat1 = layout.category(1).unwrap();
        assert_eq!(cat1.header_row, Some(26));
        assert_eq!(cat1.positions.as_ref().unwrap().start_row, 27);
        assert_eq!(cat1.positions.as_ref().unwrap().end_row, 46);
        assert_eq!(cat1.positions.as_ref().unwrap().count, 20);
        assert_eq!(cat1.footer_row, Some(47));

        // Kategorie 6 (Single-Row)
        let cat6 = layout.category(6).unwrap();
        assert!(cat6.single_row.is_some());
        assert!(cat6.positions.is_none());
    }

    #[test]
    fn test_layout_with_custom_config() {
        let config = BodyConfig::new().with_positions(1, 5).with_positions(2, 10);
        let layout = BodyLayout::compute(&config);

        // Kategorie 1: 5 Positionen
        let cat1 = layout.category(1).unwrap();
        assert_eq!(cat1.positions.as_ref().unwrap().count, 5);
        // Header=26, Pos=27-31, Footer=32
        assert_eq!(cat1.footer_row, Some(32));

        // Kategorie 2: 10 Positionen, startet bei 33
        let cat2 = layout.category(2).unwrap();
        assert_eq!(cat2.header_row, Some(33));
        assert_eq!(cat2.positions.as_ref().unwrap().count, 10);
    }

    #[test]
    fn test_footer_and_single_rows() {
        let config = BodyConfig::default();
        let layout = BodyLayout::compute(&config);

        // 5 Footer-Zeilen (Kategorien 1-5)
        assert_eq!(layout.footer_rows().len(), 5);

        // 3 Single-Rows (Kategorien 6-8)
        assert_eq!(layout.single_rows().len(), 3);
    }

    #[test]
    fn test_ratio_rows() {
        let config = BodyConfig::new()
            .with_positions(1, 2)
            .with_positions(2, 2)
            .with_positions(3, 2)
            .with_positions(4, 2)
            .with_positions(5, 2);
        let layout = BodyLayout::compute(&config);

        let ratio_rows = layout.ratio_rows();

        // 5 Kategorien × 2 Positionen = 10 Position-Zeilen
        // + 5 Footer-Zeilen
        // + 3 Single-Row-Zeilen
        // + 1 Gesamt-Zeile
        // = 19 Zeilen mit Ratio-Formel
        assert_eq!(ratio_rows.len(), 19);
    }

    #[test]
    fn test_position_addr() {
        use PositionField::*;

        let config = BodyConfig::new().with_positions(1, 5);
        let layout = BodyLayout::compute(&config);

        // Kategorie 1: Header=26, Pos=27-31, Footer=32
        // Position 1 (erste) = Row 27
        let addr = layout.position_addr(1, 1, Description).unwrap();
        assert_eq!(addr.row, 27);
        assert_eq!(addr.col, 2); // C

        // Position 1, Spalte D
        let addr = layout.position_addr(1, 1, Approved).unwrap();
        assert_eq!(addr.row, 27);
        assert_eq!(addr.col, 3); // D

        // Position 5 (letzte) = Row 31
        let addr = layout.position_addr(1, 5, IncomeTotal).unwrap();
        assert_eq!(addr.row, 31);
        assert_eq!(addr.col, 5); // F

        // Position 6 existiert nicht
        assert!(layout.position_addr(1, 6, Description).is_none());

        // Position 0 ist ungültig (1-basiert!)
        assert!(layout.position_addr(1, 0, Description).is_none());

        // Kategorie 6 (Single-Row) hat keine Positionen
        assert!(layout.position_addr(6, 1, Description).is_none());

        // Kategorie 9 existiert nicht
        assert!(layout.position_addr(9, 1, Description).is_none());
    }
}
