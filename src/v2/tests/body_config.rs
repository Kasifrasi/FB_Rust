//! Tests für BodyConfig
//!
//! Testet die flexible Kategorien-Konfiguration:
//! - 0 Positionen = Header-Eingabe
//! - 1+ Positionen = Positions-Modus

use crate::v2::report::BodyConfig;

#[test]
fn test_default_config() {
    let config = BodyConfig::default();

    // Kategorien 1-5: Standard-Positionen
    assert_eq!(config.position_count(1), 20);
    assert_eq!(config.position_count(2), 20);
    assert_eq!(config.position_count(3), 30);
    assert_eq!(config.position_count(4), 30);
    assert_eq!(config.position_count(5), 20);

    // Kategorien 6-8: Header-Eingabe (0 Positionen)
    assert_eq!(config.position_count(6), 0);
    assert_eq!(config.position_count(7), 0);
    assert_eq!(config.position_count(8), 0);
}

#[test]
fn test_custom_positions() {
    let config = BodyConfig::new()
        .with_positions(1, 10)
        .with_positions(2, 5)
        .with_positions(3, 15);

    assert_eq!(config.position_count(1), 10);
    assert_eq!(config.position_count(2), 5);
    assert_eq!(config.position_count(3), 15);

    // Nicht explizit gesetzte Kategorien haben Default-Werte aus BodyConfig::default()
    // Kategorie 4: Default ist 30
    // Kategorie 6: Default ist 0 (Header-Eingabe)
    assert_eq!(config.position_count(4), 30);
    assert_eq!(config.position_count(6), 0);
}

#[test]
fn test_header_input_mode() {
    let config = BodyConfig::new()
        .with_positions(1, 0) // Header-Eingabe
        .with_positions(2, 1) // 1 Position
        .with_positions(3, 5); // 5 Positionen

    assert!(config.is_header_input(1));
    assert!(!config.is_header_input(2));
    assert!(!config.is_header_input(3));

    assert!(!config.has_positions(1));
    assert!(config.has_positions(2));
    assert!(config.has_positions(3));
}

#[test]
fn test_all_categories_flexible() {
    // Jede Kategorie kann zwischen Header-Eingabe und Positions-Modus wechseln
    let config = BodyConfig::new()
        .with_positions(1, 0) // Normalerweise Multi-Row, jetzt Header-Eingabe
        .with_positions(2, 0) // Header-Eingabe
        .with_positions(3, 1) // 1 Position
        .with_positions(4, 10) // 10 Positionen
        .with_positions(5, 0) // Header-Eingabe
        .with_positions(6, 5) // Normalerweise Header-Eingabe, jetzt Multi-Row!
        .with_positions(7, 3) // Multi-Row
        .with_positions(8, 0); // Header-Eingabe (Standard)

    // Header-Eingabe
    assert!(config.is_header_input(1));
    assert!(config.is_header_input(2));
    assert!(config.is_header_input(5));
    assert!(config.is_header_input(8));

    // Positions-Modus
    assert!(config.has_positions(3));
    assert!(config.has_positions(4));
    assert!(config.has_positions(6));
    assert!(config.has_positions(7));
}

#[test]
fn test_invalid_category_ignored() {
    let config = BodyConfig::new()
        .with_positions(0, 10) // Ungültig (Kategorie 0 existiert nicht)
        .with_positions(9, 10) // Ungültig (Kategorie 9 existiert nicht)
        .with_positions(1, 5); // Gültig

    assert_eq!(config.position_count(1), 5);
    assert_eq!(config.position_count(0), 0); // Default
    assert_eq!(config.position_count(9), 0); // Default
}

#[test]
fn test_position_count_minimum() {
    // 0 ist das Minimum (Header-Eingabe)
    let config = BodyConfig::new().with_positions(1, 0).with_positions(2, 1);

    assert_eq!(config.position_count(1), 0);
    assert_eq!(config.position_count(2), 1);
}
