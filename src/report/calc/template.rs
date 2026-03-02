//! Master-Template für IronCalc-Modell
//!
//! Enthält alle statischen Formeln und die Sprachversionen-Daten.
//! Wird einmal erstellt und per `to_bytes()`/`from_bytes()` für jeden Report geklont.

use crate::lang::data::TEXT_MATRIX;
use crate::report::core::CellAddr;
use ironcalc_base::Model;

/// Sheet-Index für den Bericht (Sheet 0)
const SHEET_BERICHT: u32 = 0;
/// Sheet-Index für Sprachversionen (Sheet 1)
const SHEET_LANG: u32 = 1;

/// Cached IronCalc-Modell mit allen statischen Formeln.
///
/// Enthält:
/// - Sheet 0 "Bericht": Alle statischen Formeln (Header, Tabelle, Panel, Prebody)
/// - Sheet 1 "Sprachversionen": Vollständige TEXT_MATRIX (5 Sprachen × 66 Begriffe)
///
/// Das Template wird einmal erstellt und per `to_bytes()`/`from_bytes()`
/// für jeden Report als Basis geklont.
pub(crate) struct ModelTemplate {
    bytes: Vec<u8>,
}

impl ModelTemplate {
    /// Erstellt das Master-Template mit allen statischen Formeln.
    pub(crate) fn new() -> Self {
        let mut model =
            Model::new_empty("Bericht", "en", "UTC", "en").expect("Failed to create IronCalc model");

        // Sheet 1: Sprachversionen hinzufügen und befüllen
        model
            .add_sheet("Sprachversionen")
            .expect("Failed to add Sprachversionen sheet");
        populate_language_sheet(&mut model);

        // Sheet 0: Alle statischen Formeln registrieren
        register_static_formulas(&mut model);

        let bytes = model.to_bytes();
        Self { bytes }
    }

    /// Erstellt ein frisches Model aus dem Template.
    pub(crate) fn instantiate(&self) -> Model<'static> {
        Model::from_bytes(&self.bytes, "en").expect("Failed to deserialize template")
    }
}

// ============================================================================
// Sprachversionen Sheet
// ============================================================================

/// Befüllt Sheet "Sprachversionen" mit TEXT_MATRIX Daten.
///
/// TEXT_MATRIX[lang_row][col_idx] → IronCalc Sheet 1, Row (lang_row+1), Col (col_idx+2)
/// - col_idx=0 → Col 2 (B) = Sprachname
/// - col_idx=1 → Col 3 (C) = erster Text (VLOOKUP index 2)
/// - col_idx=N → Col (N+2)
fn populate_language_sheet(model: &mut Model) {
    for (row_idx, row) in TEXT_MATRIX.iter().enumerate() {
        let ic_row = row_idx as i32 + 1; // 1-basiert
        for (col_idx, text) in row.iter().enumerate() {
            let ic_col = col_idx as i32 + 2; // B=2, C=3, ...
            model
                .set_user_input(SHEET_LANG, ic_row, ic_col, text.to_string())
                .expect("Failed to set language cell");
        }
    }
}

// ============================================================================
// Statische Formeln
// ============================================================================

/// Registriert alle statischen Formeln auf Sheet "Bericht".
///
/// Übersetzt definitions.rs `register_formula_cells()` + `register_prebody_formulas()`.
fn register_static_formulas(model: &mut Model) {
    // ========================================================================
    // VLOOKUP Text-Lookups
    // ========================================================================

    // Row 0: B1 (Titel), J1 (Ausfüllbare Felder Info)
    set_vlookup(model, 0, 1, 2); // B1
    set_vlookup(model, 0, 9, 60); // J1

    // Row 1: D2 (Sprache Label mit Default), J2 (Tipp)
    set_vlookup_default(model, 1, 3, 27, "Chose your language"); // D2
    set_vlookup(model, 1, 9, 61); // J2

    // Row 2: B3 (Übersicht), D3 (Lokalwährung Label)
    set_vlookup(model, 2, 1, 3); // B3
    set_vlookup(model, 2, 3, 28); // D3

    // Row 3: J4 (Hyperlink)
    set_hyperlink(model, 3, 9, 62); // J4

    // Row 4: B5 (Projektnummer Label)
    set_vlookup(model, 4, 1, 4); // B5

    // Row 5: B6 (Projekttitel Label)
    set_vlookup(model, 5, 1, 26); // B6

    // Row 7: B8 (Projektlaufzeit), D8 (von), F8 (bis)
    set_vlookup(model, 7, 1, 5); // B8
    set_vlookup(model, 7, 3, 7); // D8
    set_vlookup(model, 7, 5, 8); // F8

    // Row 8: B9 (Berichtszeitraum), D9 (von), F9 (bis), K8 (WK Label)
    set_vlookup(model, 8, 1, 6); // B9
    set_vlookup(model, 8, 3, 7); // D9
    set_vlookup(model, 8, 5, 8); // F9
    set_vlookup(model, 7, 10, 59); // K8

    // Row 10: Tabellenkopf
    set_vlookup(model, 10, 9, 19); // J11
    set_vlookup(model, 10, 16, 19); // Q11
    set_vlookup(model, 10, 3, 11); // D11
    set_vlookup(model, 10, 4, 12); // E11
    set_vlookup(model, 10, 5, 13); // F11
    set_vlookup(model, 10, 6, 14); // G11
    set_vlookup(model, 10, 7, 15); // H11

    // Row 11: B12 (EINNAHMEN)
    set_vlookup(model, 11, 1, 9); // B12

    // Row 12: B13, Panel Headers
    set_vlookup(model, 12, 1, 10); // B13
    set_vlookup(model, 12, 11, 22); // L13: Datum
    set_vlookup(model, 12, 12, 63); // M13: Euro
    set_currency_or_lookup(model, 12, 13, 28); // N13: Währung
    set_vlookup(model, 12, 14, 58); // O13: Wechselkurs
    set_vlookup(model, 12, 18, 22); // S13: Datum
    set_vlookup(model, 12, 19, 63); // T13: Euro
    set_currency_or_lookup(model, 12, 20, 28); // U13: Währung
    set_vlookup(model, 12, 21, 58); // V13: Wechselkurs

    // Row 14-18: B15-B19 (Tabellen-Labels)
    set_vlookup(model, 14, 1, 16); // B15: Saldovortrag
    set_vlookup(model, 15, 1, 17); // B16: Lokale Eigenleistung
    set_vlookup(model, 16, 1, 18); // B17: Beiträge dritter Seite
    set_vlookup(model, 17, 1, 19); // B18: Zuschuss-Mittel
    set_vlookup(model, 18, 1, 20); // B19: Zinserträge

    // ========================================================================
    // IFERROR Division (G16-G19)
    // ========================================================================

    for i in 1..5u32 {
        let row = 14 + i; // rows 15-18
        set_iferror_division(model, row, 6, row, 5, row, 3); // G = IFERROR(F/D, 0)
    }

    // ========================================================================
    // Row 19: B20 (GESAMT), D20-G20
    // ========================================================================

    set_vlookup(model, 19, 1, 21); // B20

    // D20: =SUM(D15:D19)
    set_sum_range(model, 19, 3, 14, 18, 3); // D20
    // E20: =SUM(E15:E19)
    set_sum_range(model, 19, 4, 14, 18, 4); // E20
    // F20: =SUM(F15:F19)
    set_sum_range(model, 19, 5, 14, 18, 5); // F20
    // G20: =IFERROR(F20/D20,0) — depends on formula cells
    set_iferror_division(model, 19, 6, 19, 5, 19, 3); // G20

    // ========================================================================
    // Right Panel Body (Rows 13-30)
    // ========================================================================

    for i in 0..18u32 {
        let row = 13 + i;
        let excel_row = row + 1;

        // K-Spalte (Col 10): VLOOKUP(23)
        set_vlookup(model, row, 10, 23);

        // O-Spalte (Col 14): =IF(M{row}="","",N{row}/M{row})
        set_formula(
            model,
            row,
            14,
            format!(r#"=IF(M{}="","",N{}/M{})"#, excel_row, excel_row, excel_row),
        );

        // R-Spalte (Col 17): VLOOKUP(23)
        set_vlookup(model, row, 17, 23);

        // V-Spalte (Col 21): =IF(T{row}="","",U{row}/T{row})
        set_formula(
            model,
            row,
            21,
            format!(r#"=IF(T{}="","",U{}/T{})"#, excel_row, excel_row, excel_row),
        );
    }

    // ========================================================================
    // Pre-Body Formeln (Rows 22-24)
    // ========================================================================

    set_vlookup(model, 22, 3, 11); // D23
    set_vlookup(model, 22, 4, 25); // E23
    set_vlookup(model, 22, 5, 55); // F23
    set_vlookup(model, 22, 6, 56); // G23
    set_vlookup(model, 22, 7, 15); // H23
    set_vlookup(model, 23, 1, 24); // B24
    set_vlookup(model, 24, 1, 10); // B25
}

// ============================================================================
// Formula Helpers
// ============================================================================

/// Setzt eine VLOOKUP-Formel: =IF($E$2="","",VLOOKUP($E$2,Sprachversionen!$B:$CD,{index},FALSE))
fn set_vlookup(model: &mut Model, row: u32, col: u16, index: usize) {
    let formula = format!(
        r#"=IF($E$2="","",VLOOKUP($E$2,Sprachversionen!$B:$CD,{},FALSE))"#,
        index
    );
    set_formula(model, row, col, formula);
}

/// Setzt eine VLOOKUP-Formel mit Default-Wert
fn set_vlookup_default(model: &mut Model, row: u32, col: u16, index: usize, default: &str) {
    let formula = format!(
        r#"=IF($E$2="","{}",VLOOKUP($E$2,Sprachversionen!$B:$CD,{},FALSE))"#,
        default, index
    );
    set_formula(model, row, col, formula);
}

/// Setzt eine HYPERLINK-VLOOKUP-Formel
fn set_hyperlink(model: &mut Model, row: u32, col: u16, index: usize) {
    let formula = format!(
        r#"=HYPERLINK(VLOOKUP($E$2,Sprachversionen!$B:$CD,{},FALSE))"#,
        index
    );
    set_formula(model, row, col, formula);
}

/// Setzt eine Currency-or-Lookup Formel: =IF(E3="",VLOOKUP(...),E3)
fn set_currency_or_lookup(model: &mut Model, row: u32, col: u16, _index: usize) {
    let formula =
        r#"=IF(E3="",VLOOKUP($E$2,Sprachversionen!$B:$CD,28,FALSE),E3)"#.to_string();
    set_formula(model, row, col, formula);
}

/// Setzt eine IFERROR Division Formel: =IFERROR(num/denom,0)
fn set_iferror_division(
    model: &mut Model,
    target_row: u32,
    target_col: u16,
    num_row: u32,
    num_col: u16,
    denom_row: u32,
    denom_col: u16,
) {
    let num = CellAddr::new(num_row, num_col).to_excel();
    let denom = CellAddr::new(denom_row, denom_col).to_excel();
    let formula = format!("=IFERROR({}/{},0)", num, denom);
    set_formula(model, target_row, target_col, formula);
}

/// Setzt eine SUM-Formel für einen zusammenhängenden Bereich
fn set_sum_range(
    model: &mut Model,
    target_row: u32,
    target_col: u16,
    start_row: u32,
    end_row: u32,
    col: u16,
) {
    let start = CellAddr::new(start_row, col).to_excel();
    let end = CellAddr::new(end_row, col).to_excel();
    let formula = format!("=SUM({}:{})", start, end);
    set_formula(model, target_row, target_col, formula);
}

/// Low-level: Setzt eine Formel in IronCalc (konvertiert 0-basiert → 1-basiert)
fn set_formula(model: &mut Model, row: u32, col: u16, formula: String) {
    model
        .update_cell_with_formula(SHEET_BERICHT, row as i32 + 1, col as i32 + 1, formula)
        .unwrap_or_else(|e| panic!("Failed to set formula at ({}, {}): {}", row, col, e));
}

// ============================================================================
// Tests
// ============================================================================

#[cfg(test)]
mod tests {
    use super::*;
    use ironcalc_base::cell::CellValue;

    #[test]
    fn test_template_creation() {
        let template = ModelTemplate::new();
        assert!(!template.bytes.is_empty());
    }

    #[test]
    fn test_template_instantiation() {
        let template = ModelTemplate::new();
        let mut model = template.instantiate();

        // Sprache setzen und evaluieren
        model
            .set_user_input(SHEET_BERICHT, 2, 5, "deutsch".to_string())
            .unwrap(); // E2
        model.evaluate();

        // B1 VLOOKUP sollte den Titel liefern
        let result = model.get_cell_value_by_index(SHEET_BERICHT, 1, 2).unwrap();
        assert_eq!(
            result,
            CellValue::String("F I N A N Z B E R I C H T".to_string()),
            "B1 VLOOKUP should return title for 'deutsch'"
        );
    }

    #[test]
    fn test_template_vlookup_english() {
        let template = ModelTemplate::new();
        let mut model = template.instantiate();

        model
            .set_user_input(SHEET_BERICHT, 2, 5, "english".to_string())
            .unwrap();
        model.evaluate();

        let result = model.get_cell_value_by_index(SHEET_BERICHT, 1, 2).unwrap();
        assert_eq!(
            result,
            CellValue::String("F I N A N C I A L  R E P O R T ".to_string()),
            "B1 VLOOKUP should return title for 'english' (trailing space in data)"
        );
    }

    #[test]
    fn test_template_default_language_label() {
        let template = ModelTemplate::new();
        let mut model = template.instantiate();

        // Keine Sprache gesetzt → Default "Chose your language"
        model.evaluate();

        let result = model.get_cell_value_by_index(SHEET_BERICHT, 2, 4).unwrap();
        assert_eq!(
            result,
            CellValue::String("Chose your language".to_string()),
            "D2 should show default when no language set"
        );
    }

    #[test]
    fn test_template_sum_and_ratio() {
        let template = ModelTemplate::new();
        let mut model = template.instantiate();

        // D15-D19: Werte setzen
        model
            .set_user_input(SHEET_BERICHT, 15, 4, "100".to_string())
            .unwrap(); // D15
        model
            .set_user_input(SHEET_BERICHT, 16, 4, "200".to_string())
            .unwrap(); // D16
        model
            .set_user_input(SHEET_BERICHT, 17, 4, "300".to_string())
            .unwrap(); // D17
        // F16: Wert setzen für Ratio
        model
            .set_user_input(SHEET_BERICHT, 16, 6, "150".to_string())
            .unwrap(); // F16

        model.evaluate();

        // D20 = SUM(D15:D19) = 100 + 200 + 300 = 600
        let d20 = model.get_cell_value_by_index(SHEET_BERICHT, 20, 4).unwrap();
        assert_eq!(d20, CellValue::Number(600.0));

        // G16 = IFERROR(F16/D16, 0) = 150/200 = 0.75
        let g16 = model.get_cell_value_by_index(SHEET_BERICHT, 16, 7).unwrap();
        assert_eq!(g16, CellValue::Number(0.75));
    }

    #[test]
    fn test_template_clone_isolation() {
        let template = ModelTemplate::new();

        // Zwei separate Instanzen
        let mut model1 = template.instantiate();
        let mut model2 = template.instantiate();

        // Verschiedene Sprachen
        model1
            .set_user_input(SHEET_BERICHT, 2, 5, "deutsch".to_string())
            .unwrap();
        model2
            .set_user_input(SHEET_BERICHT, 2, 5, "english".to_string())
            .unwrap();

        model1.evaluate();
        model2.evaluate();

        let b1_de = model1
            .get_cell_value_by_index(SHEET_BERICHT, 1, 2)
            .unwrap();
        let b1_en = model2
            .get_cell_value_by_index(SHEET_BERICHT, 1, 2)
            .unwrap();

        assert_ne!(b1_de, b1_en, "Templates should be isolated");
    }

    #[test]
    fn test_currency_or_lookup() {
        let template = ModelTemplate::new();
        let mut model = template.instantiate();

        // Keine Währung gesetzt, Sprache "deutsch" → VLOOKUP(28) = Währungstext
        model
            .set_user_input(SHEET_BERICHT, 2, 5, "deutsch".to_string())
            .unwrap();
        model.evaluate();

        let n13_no_currency = model
            .get_cell_value_by_index(SHEET_BERICHT, 13, 14)
            .unwrap();
        // Sollte den Währungstext aus VLOOKUP(28) zurückgeben
        assert!(
            !matches!(n13_no_currency, CellValue::None),
            "N13 should have a value when no currency set"
        );

        // Jetzt mit Währung "USD" → E3 wird direkt verwendet
        let mut model2 = template.instantiate();
        model2
            .set_user_input(SHEET_BERICHT, 2, 5, "deutsch".to_string())
            .unwrap();
        model2
            .set_user_input(SHEET_BERICHT, 3, 5, "USD".to_string())
            .unwrap();
        model2.evaluate();

        let n13_with_currency = model2
            .get_cell_value_by_index(SHEET_BERICHT, 13, 14)
            .unwrap();
        assert_eq!(
            n13_with_currency,
            CellValue::String("USD".to_string()),
            "N13 should show currency when E3 is set"
        );
    }
}
