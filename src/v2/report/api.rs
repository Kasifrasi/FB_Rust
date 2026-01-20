//! API-Zellen Definition
//!
//! Alle API-Zellen werden hier zentral definiert. Das Macro generiert:
//! - `ApiKey` enum mit allen Varianten
//! - `addr()` Methode für jede Variante
//! - `register_all_api_cells()` Funktion
//!
//! Um eine neue API-Zelle hinzuzufügen:
//! 1. Einzelzelle: `NewField => (row, col),` im `single` Block
//! 2. Bereichszelle: `NewField => (start_row, col, count),` im `range` Block
//! FERTIG! Alles andere wird automatisch generiert.

use super::registry::{CellAddr, CellRegistry, RegistryError};

/// Macro zur Definition aller API-Zellen
///
/// Generiert:
/// - `ApiKey` enum
/// - `ApiKey::addr()` Methode
/// - `register_all_api_cells()` Funktion
/// - `ApiKey::all_keys()` Iterator
macro_rules! define_api_cells {
    (
        single {
            $( $single_name:ident => ($single_row:expr, $single_col:expr) ),* $(,)?
        }
        range {
            $( $range_name:ident => ($range_start_row:expr, $range_col:expr, $range_count:expr) ),* $(,)?
        }
    ) => {
        /// Schlüssel für API-Werte
        ///
        /// Jede Variante entspricht genau einer Eingabezelle im Finanzbericht.
        /// Die Adressen sind 0-basiert (row, col).
        #[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
        pub enum ApiKey {
            // Einzelzellen
            $( $single_name, )*
            // Bereichszellen mit Index
            $( $range_name(u8), )*
        }

        impl ApiKey {
            /// Gibt die Zelladresse zurück (0-basiert)
            pub const fn addr(&self) -> CellAddr {
                match self {
                    // Einzelzellen
                    $( Self::$single_name => CellAddr::new($single_row, $single_col), )*
                    // Bereichszellen
                    $( Self::$range_name(i) => CellAddr::new($range_start_row + *i as u32, $range_col), )*
                }
            }

            /// Gibt alle API-Keys zurück (für Iteration)
            pub fn all_keys() -> impl Iterator<Item = ApiKey> {
                let singles = [
                    $( Self::$single_name, )*
                ].into_iter();

                let ranges = [
                    $( (Self::$range_name as fn(u8) -> ApiKey, $range_count as u8), )*
                ].into_iter().flat_map(|(constructor, count)| {
                    (0..count).map(move |i| constructor(i))
                });

                singles.chain(ranges)
            }

            /// Anzahl aller API-Zellen
            pub const fn count() -> usize {
                // Einzelzellen zählen
                let singles = [$( stringify!($single_name), )*].len();
                // Bereichszellen summieren
                let ranges = 0 $( + $range_count )*;
                singles + ranges
            }
        }

        /// Registriert alle API-Zellen in der Registry
        pub fn register_all_api_cells<E>(
            registry: &mut CellRegistry<E>
        ) -> Result<(), RegistryError> {
            // Einzelzellen registrieren
            $( registry.register_api(ApiKey::$single_name)?; )*

            // Bereichszellen registrieren
            $(
                for i in 0..$range_count as u8 {
                    registry.register_api(ApiKey::$range_name(i))?;
                }
            )*

            Ok(())
        }
    };
}

// =============================================================================
// API-Zellen Definition - EINZIGE QUELLE DER WAHRHEIT
// =============================================================================
//
// Um eine neue Zelle hinzuzufügen:
// - Einzelzelle: Name => (row, col),
// - Bereichszelle: Name => (start_row, col, count),
//
// row/col sind 0-basiert! E2 = (1, 4), D15 = (14, 3), L14 = (13, 11)

define_api_cells! {
    single {
        // Header-Bereich
        Language => (1, 4),        // E2: Sprachauswahl
        Currency => (2, 4),        // E3: Währungsauswahl
        ProjectNumber => (4, 3),   // D5: Projektnummer
        ProjectTitle => (5, 3),    // D6: Projekttitel
        ProjectStart => (7, 4),    // E8: Projektstart
        ProjectEnd => (7, 6),      // G8: Projektende
        ReportStart => (8, 4),     // E9: Berichtszeitraum Start
        ReportEnd => (8, 6),       // G9: Berichtszeitraum Ende
    }
    range {
        // Tabelle: Zeilen 15-19 (Index 0-4)
        ApprovedBudget => (14, 3, 5),      // D15-D19: Bewilligtes Budget
        IncomeReportPeriod => (14, 4, 5),  // E15-E19: Einnahmen Berichtszeitraum
        IncomeTotal => (14, 5, 5),         // F15-F19: Einnahmen gesamt
        IncomeReason => (14, 7, 5),        // H15-H19: Begründung

        // Right Panel Links: Zeilen 14-31 (Index 0-17)
        LeftDate => (13, 11, 18),          // L14-L31: Datum
        LeftAmountEuro => (13, 12, 18),    // M14-M31: Betrag Euro
        LeftAmountLocal => (13, 13, 18),   // N14-N31: Betrag Lokal

        // Right Panel Rechts: Zeilen 14-31 (Index 0-17)
        RightDate => (13, 18, 18),         // S14-S31: Datum
        RightAmountEuro => (13, 19, 18),   // T14-T31: Betrag Euro
        RightAmountLocal => (13, 20, 18),  // U14-U31: Betrag Lokal
    }
}

// =============================================================================
// Tests
// =============================================================================

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_single_cell_addresses() {
        assert_eq!(ApiKey::Language.addr(), CellAddr::new(1, 4)); // E2
        assert_eq!(ApiKey::Currency.addr(), CellAddr::new(2, 4)); // E3
        assert_eq!(ApiKey::ProjectNumber.addr(), CellAddr::new(4, 3)); // D5
        assert_eq!(ApiKey::ProjectStart.addr(), CellAddr::new(7, 4)); // E8
        assert_eq!(ApiKey::ReportEnd.addr(), CellAddr::new(8, 6)); // G9
    }

    #[test]
    fn test_range_cell_addresses() {
        // ApprovedBudget: D15-D19
        assert_eq!(ApiKey::ApprovedBudget(0).addr(), CellAddr::new(14, 3)); // D15
        assert_eq!(ApiKey::ApprovedBudget(4).addr(), CellAddr::new(18, 3)); // D19

        // LeftDate: L14-L31
        assert_eq!(ApiKey::LeftDate(0).addr(), CellAddr::new(13, 11)); // L14
        assert_eq!(ApiKey::LeftDate(17).addr(), CellAddr::new(30, 11)); // L31

        // RightAmountLocal: U14-U31
        assert_eq!(ApiKey::RightAmountLocal(0).addr(), CellAddr::new(13, 20)); // U14
        assert_eq!(ApiKey::RightAmountLocal(17).addr(), CellAddr::new(30, 20)); // U31
    }

    #[test]
    fn test_all_keys_count() {
        let count = ApiKey::all_keys().count();
        // 8 Einzelzellen + 4*5 Table + 6*18 Right Panel = 8 + 20 + 108 = 136
        assert_eq!(count, 136);
        assert_eq!(ApiKey::count(), 136);
    }

    #[test]
    fn test_all_keys_iteration() {
        let keys: Vec<_> = ApiKey::all_keys().collect();

        // Erste 8 sind Einzelzellen
        assert_eq!(keys[0], ApiKey::Language);
        assert_eq!(keys[7], ApiKey::ReportEnd);

        // Dann kommen die Bereichszellen
        assert_eq!(keys[8], ApiKey::ApprovedBudget(0));
        assert_eq!(keys[12], ApiKey::ApprovedBudget(4));
    }
}
