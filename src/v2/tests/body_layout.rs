//! Tests für BodyLayout
//!
//! Testet die Layout-Berechnung für verschiedene Konfigurationen.

use crate::v2::report::body::CategoryMode;
use crate::v2::report::{BodyConfig, BodyLayout};

#[test]
fn test_layout_default() {
    let config = BodyConfig::default();
    let layout = BodyLayout::compute(&config);

    // 8 Kategorien
    assert_eq!(layout.categories.len(), 8);

    // Kategorien 1-5: Positions-Modus
    for i in 1..=5 {
        let cat = layout.category(i).unwrap();
        assert!(
            cat.has_positions(),
            "Kategorie {} sollte Positionen haben",
            i
        );
    }

    // Kategorien 6-8: Header-Eingabe
    for i in 6..=8 {
        let cat = layout.category(i).unwrap();
        assert!(
            cat.is_header_input(),
            "Kategorie {} sollte Header-Eingabe sein",
            i
        );
    }
}

#[test]
fn test_layout_header_input_mode() {
    let config = BodyConfig::new()
        .with_positions(1, 0) // Header-Eingabe
        .with_positions(2, 0) // Header-Eingabe
        .with_positions(3, 0); // Header-Eingabe

    let layout = BodyLayout::compute(&config);

    // Kategorie 1: Header-Eingabe bei Row 26
    let cat1 = layout.category(1).unwrap();
    assert!(cat1.is_header_input());
    if let CategoryMode::HeaderInput { row } = cat1.mode {
        assert_eq!(row, 26);
    } else {
        panic!("Erwartet HeaderInput");
    }

    // Kategorie 2: Header-Eingabe bei Row 27
    let cat2 = layout.category(2).unwrap();
    assert!(cat2.is_header_input());
    if let CategoryMode::HeaderInput { row } = cat2.mode {
        assert_eq!(row, 27);
    } else {
        panic!("Erwartet HeaderInput");
    }
}

#[test]
fn test_layout_positions_mode() {
    let config = BodyConfig::new().with_positions(1, 5).with_positions(2, 3);

    let layout = BodyLayout::compute(&config);

    // Kategorie 1: 5 Positionen
    let cat1 = layout.category(1).unwrap();
    assert!(cat1.has_positions());
    if let CategoryMode::WithPositions {
        header_row,
        positions,
        footer_row,
    } = &cat1.mode
    {
        assert_eq!(*header_row, 26);
        assert_eq!(positions.start_row, 27);
        assert_eq!(positions.end_row, 31);
        assert_eq!(positions.count, 5);
        assert_eq!(*footer_row, 32);
    } else {
        panic!("Erwartet WithPositions");
    }

    // Kategorie 2: 3 Positionen, startet nach Kategorie 1
    let cat2 = layout.category(2).unwrap();
    if let CategoryMode::WithPositions {
        header_row,
        positions,
        footer_row,
    } = &cat2.mode
    {
        assert_eq!(*header_row, 33);
        assert_eq!(positions.start_row, 34);
        assert_eq!(positions.end_row, 36);
        assert_eq!(positions.count, 3);
        assert_eq!(*footer_row, 37);
    } else {
        panic!("Erwartet WithPositions");
    }
}

#[test]
fn test_layout_mixed_modes() {
    let config = BodyConfig::new()
        .with_positions(1, 2) // Positions-Modus
        .with_positions(2, 0) // Header-Eingabe
        .with_positions(3, 3); // Positions-Modus

    let layout = BodyLayout::compute(&config);

    // Kategorie 1: Header=26, Pos=27-28, Footer=29
    let cat1 = layout.category(1).unwrap();
    assert!(cat1.has_positions());

    // Kategorie 2: Header-Eingabe bei Row 30
    let cat2 = layout.category(2).unwrap();
    assert!(cat2.is_header_input());
    assert_eq!(cat2.header_input_row(), Some(30));

    // Kategorie 3: Header=31, Pos=32-34, Footer=35
    let cat3 = layout.category(3).unwrap();
    assert!(cat3.has_positions());
    assert_eq!(cat3.header_row(), Some(31));
}

#[test]
fn test_sum_row() {
    let config = BodyConfig::new()
        .with_positions(1, 2) // Footer ist Summen-Zeile
        .with_positions(2, 0); // Header-Input ist Summen-Zeile

    let layout = BodyLayout::compute(&config);

    // Kategorie 1: sum_row = footer_row
    let cat1 = layout.category(1).unwrap();
    assert_eq!(cat1.sum_row(), cat1.footer_row().unwrap());

    // Kategorie 2: sum_row = header_input_row
    let cat2 = layout.category(2).unwrap();
    assert_eq!(cat2.sum_row(), cat2.header_input_row().unwrap());
}

#[test]
fn test_ratio_rows() {
    let config = BodyConfig::new()
        .with_positions(1, 2) // 2 Positions + 1 Footer = 3
        .with_positions(2, 0) // 1 Header-Input
        .with_positions(3, 0) // 1 Header-Input
        .with_positions(4, 1) // 1 Position + 1 Footer = 2
        .with_positions(5, 0) // 1 Header-Input
        .with_positions(6, 0) // 1 Header-Input
        .with_positions(7, 0) // 1 Header-Input
        .with_positions(8, 0); // 1 Header-Input

    let layout = BodyLayout::compute(&config);
    let ratio_rows = layout.ratio_rows();

    // 3 + 1 + 1 + 2 + 1 + 1 + 1 + 1 + 1 (Total) = 12
    assert_eq!(ratio_rows.len(), 12);
}

#[test]
fn test_total_row() {
    let config = BodyConfig::new()
        .with_positions(1, 2) // 1 Header + 2 Pos + 1 Footer = 4 Zeilen
        .with_positions(2, 0) // 1 Header-Input
        .with_positions(3, 0) // 1 Header-Input
        .with_positions(4, 0) // 1 Header-Input
        .with_positions(5, 0) // 1 Header-Input
        .with_positions(6, 0) // 1 Header-Input
        .with_positions(7, 0) // 1 Header-Input
        .with_positions(8, 0); // 1 Header-Input

    let layout = BodyLayout::compute(&config);

    // Start bei 26 + 4 (Kat1) + 7 (Kat2-8 Header-Input) = 37
    assert_eq!(layout.total_row, 37);
}

#[test]
fn test_footer_and_header_input_rows() {
    let config = BodyConfig::new()
        .with_positions(1, 2)
        .with_positions(2, 3)
        .with_positions(3, 0)
        .with_positions(4, 0)
        .with_positions(5, 1)
        .with_positions(6, 0)
        .with_positions(7, 0)
        .with_positions(8, 0);

    let layout = BodyLayout::compute(&config);

    // Footer-Zeilen: Kat 1, 2, 5
    let footer_rows = layout.footer_rows();
    assert_eq!(footer_rows.len(), 3);

    // Header-Input-Zeilen: Kat 3, 4, 6, 7, 8
    let header_input_rows = layout.header_input_rows();
    assert_eq!(header_input_rows.len(), 5);
}
