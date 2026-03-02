//! Body-Konfiguration
//!
//! Definiert die Anzahl der Positionen pro Kategorie.
//!
//! ## Konzept: Flexible Kategorien
//!
//! Jede Kategorie (1-8) kann flexibel konfiguriert werden:
//!
//! - **0 Positionen**: Header-Zeile ist die Eingabezeile (wie Single-Row)
//!   - Nur eine Zeile für die Kategorie
//!   - Eingabe direkt im Header (Spalten D, E, F, H)
//!   - Kein Footer
//!
//! - **1+ Positionen**: Positionen unter dem Header
//!   - Header-Zeile (Kategorie-Label)
//!   - N Positions-Zeilen (Eingabe)
//!   - Footer-Zeile (Summen)

use crate::report::api::Category;
use std::collections::HashMap;

/// Start-Zeile für den dynamischen Body (0-basiert)
/// Excel Row 27 = Index 26
pub const BODY_START_ROW: u32 = 26;

/// Alle Kategorien (1-8)
pub const ALL_CATEGORIES: [u8; 8] = [1, 2, 3, 4, 5, 6, 7, 8];

/// Configuration for the dynamic body area.
///
/// ## Example
/// ```ignore
/// let config = BodyConfig::new()
///     .with_positions(1, 10)   // Category 1: 10 position rows under header
///     .with_positions(2, 5)    // Category 2: 5 positions
///     .with_positions(6, 0)    // Category 6: header-input (0 positions)
///     .with_positions(7, 1);   // Category 7: 1 position under header
/// ```
///
/// ## Serde (requires `serde` feature)
///
/// Serialized as a flat JSON object via `from`/`into` `HashMap<u8, u16>`:
///
/// ```json
/// { "1": 20, "2": 20, "3": 30, "4": 30, "5": 20, "6": 0, "7": 0, "8": 0 }
/// ```
///
/// Keys are category numbers (1–8) as strings, values are the number of position
/// rows. `0` means header-input mode (single-row, no footer).
///
/// **Note:** In [`ReportConfig`](crate::ReportConfig), this is derived automatically
/// from the `positions` and `header_inputs` fields — it is not exposed in the public API.
///
/// Defaults: see [`BodyConfig::default_positions`].
#[derive(Debug, Clone, PartialEq, Eq)]
#[cfg_attr(feature = "serde", derive(serde::Serialize, serde::Deserialize))]
#[cfg_attr(feature = "serde", serde(from = "HashMap<u8, u16>", into = "HashMap<u8, u16>"))]
pub struct BodyConfig {
    /// Anzahl Positionen pro Kategorie (0 = Header-Eingabe)
    positions: HashMap<u8, u16>,
}

impl BodyConfig {
    /// Erstellt eine neue Konfiguration mit Standard-Werten
    pub fn new() -> Self {
        Self::default()
    }

    /// Setzt die Anzahl Positionen für eine Kategorie
    ///
    /// - `count = 0`: Header-Eingabe (Single-Row-Modus)
    /// - `count >= 1`: Positionen unter dem Header
    pub fn with_positions(mut self, category: u8, count: u16) -> Self {
        if ALL_CATEGORIES.contains(&category) {
            self.positions.insert(category, count);
        }
        self
    }

    /// Setzt die Anzahl Positionen typsicher mit Category enum
    ///
    /// # Beispiel
    /// ```ignore
    /// use fb_rust::report::types::Category;
    ///
    /// let config = BodyConfig::new()
    ///     .with_cat_positions(Category::Personal, 10)
    ///     .with_cat_positions(Category::Projektverwaltung, 0);  // Header-Eingabe
    /// ```
    pub fn with_cat_positions(self, category: Category, count: u16) -> Self {
        self.with_positions(category.index(), count)
    }

    /// Gibt die Anzahl Positionen für eine Kategorie zurück
    ///
    /// - `0`: Header-Eingabe (keine separaten Positions-Zeilen)
    /// - `1+`: Anzahl der Positions-Zeilen unter dem Header
    pub fn position_count(&self, category: u8) -> u16 {
        *self.positions.get(&category).unwrap_or(&0)
    }

    /// Prüft ob eine Kategorie im Header-Eingabe-Modus ist (0 Positionen)
    pub fn is_header_input(&self, category: u8) -> bool {
        self.position_count(category) == 0
    }

    /// Prüft ob eine Kategorie Positionen unter dem Header hat (1+ Positionen)
    pub fn has_positions(&self, category: u8) -> bool {
        self.position_count(category) > 0
    }
}

impl Default for BodyConfig {
    fn default() -> Self {
        Self {
            positions: Self::default_positions(),
        }
    }
}

impl BodyConfig {
    /// Gibt die Standard-Positions-Map zurück (Kategorien 1-5: 20/30 Zeilen, 6-8: Header-Eingabe)
    pub fn default_positions() -> HashMap<u8, u16> {
        HashMap::from([
            (1, 20),
            (2, 20),
            (3, 30),
            (4, 30),
            (5, 20),
            (6, 0), // Header-Eingabe
            (7, 0), // Header-Eingabe
            (8, 0), // Header-Eingabe
        ])
    }
}

#[cfg(feature = "serde")]
impl From<HashMap<u8, u16>> for BodyConfig {
    fn from(positions: HashMap<u8, u16>) -> Self {
        BodyConfig { positions }
    }
}

#[cfg(feature = "serde")]
impl From<BodyConfig> for HashMap<u8, u16> {
    fn from(config: BodyConfig) -> Self {
        config.positions
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_default_config() {
        let config = BodyConfig::default();

        // Multi-Row Kategorien
        assert_eq!(config.position_count(1), 20);
        assert_eq!(config.position_count(3), 30);
        assert!(config.has_positions(1));
        assert!(!config.is_header_input(1));

        // Header-Eingabe Kategorien (vormals Single-Row)
        assert_eq!(config.position_count(6), 0);
        assert_eq!(config.position_count(7), 0);
        assert_eq!(config.position_count(8), 0);
        assert!(!config.has_positions(6));
        assert!(config.is_header_input(6));
    }

    #[test]
    fn test_custom_config() {
        let config = BodyConfig::new().with_positions(1, 10).with_positions(6, 5); // Kategorie 6 jetzt mit Positionen!

        assert_eq!(config.position_count(1), 10);
        assert_eq!(config.position_count(6), 5);
        assert!(config.has_positions(6));
        assert!(!config.is_header_input(6));
    }

    #[test]
    fn test_header_input_mode() {
        let config = BodyConfig::new()
            .with_positions(1, 0) // Kategorie 1 im Header-Eingabe-Modus!
            .with_positions(2, 1); // Kategorie 2 mit 1 Position

        assert!(config.is_header_input(1));
        assert!(!config.is_header_input(2));
        assert_eq!(config.position_count(2), 1);
    }

    #[test]
    fn test_flexible_categories() {
        // Alle Kategorien können zwischen Header-Eingabe und Positionen wechseln
        let config = BodyConfig::new()
            .with_positions(1, 0) // Normalerweise Multi-Row, jetzt Header-Eingabe
            .with_positions(6, 10) // Normalerweise Header-Eingabe, jetzt Multi-Row
            .with_positions(7, 1) // 1 Position
            .with_positions(8, 0); // Header-Eingabe (Standard)

        assert!(config.is_header_input(1));
        assert!(config.has_positions(6));
        assert!(config.has_positions(7));
        assert!(config.is_header_input(8));
    }
}
