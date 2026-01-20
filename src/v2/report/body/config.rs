//! Body-Konfiguration
//!
//! Definiert die Anzahl der Positionen pro Kategorie.

use std::collections::HashMap;

/// Start-Zeile für den dynamischen Body (0-basiert)
/// Excel Row 27 = Index 26
pub const BODY_START_ROW: u32 = 26;

/// Multi-Row Kategorien (haben Header, N Positionen, Footer)
pub const MULTI_ROW_CATEGORIES: [u8; 5] = [1, 2, 3, 4, 5];

/// Single-Row Kategorien (nur eine Zeile, keine Expansion)
pub const SINGLE_ROW_CATEGORIES: [u8; 3] = [6, 7, 8];

/// Konfiguration für den dynamischen Body
#[derive(Debug, Clone)]
pub struct BodyConfig {
    /// Anzahl Positionen pro Kategorie (nur für Kategorien 1-5)
    pub positions: HashMap<u8, u16>,
}

impl BodyConfig {
    /// Erstellt eine neue Konfiguration
    pub fn new() -> Self {
        Self::default()
    }

    /// Setzt die Anzahl Positionen für eine Kategorie
    pub fn with_positions(mut self, category: u8, count: u16) -> Self {
        if MULTI_ROW_CATEGORIES.contains(&category) {
            self.positions.insert(category, count);
        }
        self
    }

    /// Gibt die Anzahl Positionen für eine Kategorie zurück
    pub fn position_count(&self, category: u8) -> u16 {
        if MULTI_ROW_CATEGORIES.contains(&category) {
            *self.positions.get(&category).unwrap_or(&20)
        } else {
            0 // Single-Row Kategorien haben keine Positionen
        }
    }
}

impl Default for BodyConfig {
    fn default() -> Self {
        Self {
            positions: HashMap::from([(1, 20), (2, 20), (3, 30), (4, 30), (5, 20)]),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_default_config() {
        let config = BodyConfig::default();
        assert_eq!(config.position_count(1), 20);
        assert_eq!(config.position_count(3), 30);
        assert_eq!(config.position_count(6), 0); // Single-Row
    }

    #[test]
    fn test_custom_config() {
        let config = BodyConfig::new()
            .with_positions(1, 10)
            .with_positions(2, 15);

        assert_eq!(config.position_count(1), 10);
        assert_eq!(config.position_count(2), 15);
        assert_eq!(config.position_count(3), 30); // Default für Kategorie 3
    }
}
