//! Tests für die Position-API
//!
//! Testet:
//! - ApiKey::Position mit position=0 (Header-Eingabe) und position>=1 (Positionen)
//! - PositionField
//! - ReportValues Convenience-Methoden
//! - BodyLayout::position_addr()

use crate::report::api::PositionField;
use crate::report::{ApiKey, BodyConfig, BodyLayout, ReportValues};

// ============================================================================
// PositionField Tests
// ============================================================================

#[test]
fn test_position_field_columns() {
    use PositionField::*;

    assert_eq!(Description.col(), 2); // C
    assert_eq!(Approved.col(), 3); // D
    assert_eq!(IncomeReport.col(), 4); // E
    assert_eq!(IncomeTotal.col(), 5); // F
    assert_eq!(Remark.col(), 7); // H
}

#[test]
fn test_position_field_all() {
    let all = PositionField::all();
    assert_eq!(all.len(), 5);
}

#[test]
fn test_position_field_header_input_fields() {
    let fields = PositionField::header_input_fields();
    assert_eq!(fields.len(), 4);

    // Description nicht enthalten
    assert!(!fields.contains(&PositionField::Description));

    // Alle anderen enthalten
    assert!(fields.contains(&PositionField::Approved));
    assert!(fields.contains(&PositionField::IncomeReport));
    assert!(fields.contains(&PositionField::IncomeTotal));
    assert!(fields.contains(&PositionField::Remark));
}

#[test]
fn test_position_field_availability() {
    use PositionField::*;

    // Description nicht bei Header-Eingabe verfügbar
    assert!(!Description.available_at_header_input());

    // Alle anderen verfügbar
    assert!(Approved.available_at_header_input());
    assert!(IncomeReport.available_at_header_input());
    assert!(IncomeTotal.available_at_header_input());
    assert!(Remark.available_at_header_input());
}

// ============================================================================
// ApiKey::Position Tests
// ============================================================================

#[test]
fn test_api_key_position_is_dynamic() {
    let key = ApiKey::Position {
        category: 1,
        position: 1,
        field: PositionField::Description,
    };

    assert!(key.is_dynamic());
    assert_eq!(key.static_addr(), None);
}

#[test]
fn test_api_key_position_header_input() {
    // position=0 für Header-Eingabe
    let key = ApiKey::Position {
        category: 6,
        position: 0,
        field: PositionField::Approved,
    };

    assert!(key.is_dynamic());
}

// ============================================================================
// ReportValues Tests
// ============================================================================

#[test]
fn test_set_position() {
    use PositionField::*;

    let mut values = ReportValues::new();
    values.set_position(1, 1, Description, "Test");
    values.set_position(1, 1, Approved, 1000.0);

    assert_eq!(
        values.get_position(1, 1, Description).as_text(),
        Some("Test")
    );
    assert_eq!(
        values.get_position(1, 1, Approved).as_number(),
        Some(1000.0)
    );
}

#[test]
fn test_set_position_row() {
    use PositionField::*;

    let mut values = ReportValues::new();
    values.set_position_row(1, 1, "Beschreibung", 5000.0, 2500.0, 2500.0, "Bemerkung");

    assert_eq!(
        values.get_position(1, 1, Description).as_text(),
        Some("Beschreibung")
    );
    assert_eq!(
        values.get_position(1, 1, Approved).as_number(),
        Some(5000.0)
    );
    assert_eq!(
        values.get_position(1, 1, IncomeReport).as_number(),
        Some(2500.0)
    );
    assert_eq!(
        values.get_position(1, 1, IncomeTotal).as_number(),
        Some(2500.0)
    );
    assert_eq!(
        values.get_position(1, 1, Remark).as_text(),
        Some("Bemerkung")
    );
}

#[test]
fn test_set_header_input() {
    use PositionField::*;

    let mut values = ReportValues::new();
    values.set_header_input(6, 4000.0, 2000.0, 2000.0, "Notiz");

    // position=0 für Header-Eingabe
    assert_eq!(
        values.get_position(6, 0, Approved).as_number(),
        Some(4000.0)
    );
    assert_eq!(
        values.get_position(6, 0, IncomeReport).as_number(),
        Some(2000.0)
    );
    assert_eq!(
        values.get_position(6, 0, IncomeTotal).as_number(),
        Some(2000.0)
    );
    assert_eq!(values.get_position(6, 0, Remark).as_text(), Some("Notiz"));

    // Description nicht gesetzt
    assert!(values.get_position(6, 0, Description).is_empty());
}

#[test]
fn test_mixed_positions_and_header_input() {
    use PositionField::*;

    let mut values = ReportValues::new();

    // Kategorie 1: Positions-Modus (position >= 1)
    values.set_position_row(1, 1, "Personal", 5000.0, 2500.0, 2500.0, "");
    values.set_position_row(1, 2, "Material", 3000.0, 1500.0, 1500.0, "");

    // Kategorie 6: Header-Eingabe (position = 0)
    values.set_header_input(6, 4000.0, 2000.0, 2000.0, "Sonstiges");

    // Prüfe Kategorie 1
    assert_eq!(
        values.get_position(1, 1, Description).as_text(),
        Some("Personal")
    );
    assert_eq!(
        values.get_position(1, 2, Description).as_text(),
        Some("Material")
    );

    // Prüfe Kategorie 6
    assert_eq!(
        values.get_position(6, 0, Approved).as_number(),
        Some(4000.0)
    );
}

#[test]
fn test_multiple_categories() {
    use PositionField::*;

    let mut values = ReportValues::new();

    // Verschiedene Kategorien
    values.set_position_row(1, 1, "Kat1-Pos1", 1000.0, 500.0, 500.0, "");
    values.set_position_row(2, 1, "Kat2-Pos1", 2000.0, 1000.0, 1000.0, "");
    values.set_position_row(3, 1, "Kat3-Pos1", 3000.0, 1500.0, 1500.0, "");

    assert_eq!(
        values.get_position(1, 1, Description).as_text(),
        Some("Kat1-Pos1")
    );
    assert_eq!(
        values.get_position(2, 1, Description).as_text(),
        Some("Kat2-Pos1")
    );
    assert_eq!(
        values.get_position(3, 1, Description).as_text(),
        Some("Kat3-Pos1")
    );
}

// ============================================================================
// BodyLayout::position_addr Tests
// ============================================================================

#[test]
fn test_position_addr_with_positions() {
    use PositionField::*;

    let config = BodyConfig::new().with_positions(1, 5);
    let layout = BodyLayout::compute(&config);

    // position=1..5 gültig
    let addr = layout.position_addr(1, 1, Description).unwrap();
    assert_eq!(addr.row, 27);
    assert_eq!(addr.col, 2); // C

    let addr = layout.position_addr(1, 5, Approved).unwrap();
    assert_eq!(addr.row, 31);
    assert_eq!(addr.col, 3); // D

    // position=0 nicht gültig bei Positions-Modus
    assert!(layout.position_addr(1, 0, Approved).is_none());

    // position=6 nicht gültig (nur 5 Positionen)
    assert!(layout.position_addr(1, 6, Approved).is_none());
}

#[test]
fn test_position_addr_header_input() {
    use PositionField::*;

    let config = BodyConfig::new().with_positions(1, 0); // Header-Eingabe
    let layout = BodyLayout::compute(&config);

    // position=0 gültig bei Header-Eingabe
    let addr = layout.position_addr(1, 0, Approved).unwrap();
    assert_eq!(addr.row, 26);
    assert_eq!(addr.col, 3); // D

    // Description nicht gültig bei Header-Eingabe
    assert!(layout.position_addr(1, 0, Description).is_none());

    // position>0 nicht gültig bei Header-Eingabe
    assert!(layout.position_addr(1, 1, Approved).is_none());
}

#[test]
fn test_position_addr_invalid_category() {
    use PositionField::*;

    let config = BodyConfig::default();
    let layout = BodyLayout::compute(&config);

    // Kategorie 0 existiert nicht
    assert!(layout.position_addr(0, 1, Approved).is_none());

    // Kategorie 9 existiert nicht
    assert!(layout.position_addr(9, 1, Approved).is_none());
}

#[test]
fn test_position_addr_all_fields() {
    use PositionField::*;

    let config = BodyConfig::new().with_positions(1, 3);
    let layout = BodyLayout::compute(&config);

    // Alle Felder für Position 1
    let row = 27; // Erste Position

    let addr = layout.position_addr(1, 1, Description).unwrap();
    assert_eq!((addr.row, addr.col), (row, 2)); // C

    let addr = layout.position_addr(1, 1, Approved).unwrap();
    assert_eq!((addr.row, addr.col), (row, 3)); // D

    let addr = layout.position_addr(1, 1, IncomeReport).unwrap();
    assert_eq!((addr.row, addr.col), (row, 4)); // E

    let addr = layout.position_addr(1, 1, IncomeTotal).unwrap();
    assert_eq!((addr.row, addr.col), (row, 5)); // F

    let addr = layout.position_addr(1, 1, Remark).unwrap();
    assert_eq!((addr.row, addr.col), (row, 7)); // H
}
