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

use crate::report::types::Category;
use std::collections::HashMap;

/// Start-Zeile für den dynamischen Body (0-basiert)
/// Excel Row 27 = Index 26
pub const BODY_START_ROW: u32 = 26;

/// Alle Kategorien (1-8)
pub const ALL_CATEGORIES: [u8; 8] = [1, 2, 3, 4, 5, 6, 7, 8];

/// Konfiguration für den dynamischen Body
///
/// ## Beispiel
/// ```ignore
/// let config = BodyConfig::new()
///     .with_positions(1, 10)   // Kategorie 1: 10 Positionen unter Header
///     .with_positions(2, 5)    // Kategorie 2: 5 Positionen
///     .with_positions(6, 0)    // Kategorie 6: Header-Eingabe (0 Positionen)
///     .with_positions(7, 1);   // Kategorie 7: 1 Position unter Header
/// ```
#[derive(Debug, Clone)]
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
    /// use kmw_fb_rust::report::types::Category;
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
        // Standard: Kategorien 1-5 haben 20/30 Positionen, 6-8 haben 0 (Header-Eingabe)
        Self {
            positions: HashMap::from([
                (1, 20),
                (2, 20),
                (3, 30),
                (4, 30),
                (5, 20),
                (6, 0), // Header-Eingabe
                (7, 0), // Header-Eingabe
                (8, 0), // Header-Eingabe
            ]),
        }
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
