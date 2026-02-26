//! Report Footer Sektion
//!
//! Der Footer kommt nach dem dynamischen Body-Bereich mit 3 Zeilen Abstand.
//! Basiert auf Template Zeilen 164-184.
//!
//! Struktur (21 Zeilen):
//! - Zeile 0: E: "Saldo für den Berichtszeitraum"
//! - Zeile 1: B: "ABSCHLUSS"
//! - Zeile 2: E: Währungs-Referenz (=B13)
//! - Zeile 3: Leer
//! - Zeile 4: B: "Saldo...", D: Check ✓, E: Differenz-Formel
//! - Zeile 5: Leer
//! - Zeile 6: B: "Saldenabstimmung:", E: OK-Check
//! - Zeile 7-9: B: "Bank/Kasse/Sonstiges", E: INPUT (API)
//! - Zeile 10-12: Leer
//! - Zeile 13: B: Bestätigung 1
//! - Zeile 14: B: Bestätigung 2
//! - Zeile 15-18: Leer
//! - Zeile 19: B: "Ort, Datum...", D: "Unterschrift..."
//! - Zeile 20: D: "Funktion..."

use rust_xlsxwriter::{Formula, Worksheet, XlsxError};

use crate::report::core::lookup_text_string;
use crate::report::format::FormatMatrix;

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

/// Schreibt den Report-Footer
///
/// # Arguments
/// * `ws` - Das Worksheet
/// * `fmt` - FormatMatrix
/// * `total_row` - Die Total-Zeile des Body (0-indexed)
/// * `income_row` - Die Einnahmen-Zeile (Zeile 20, 0-indexed = 19)
/// * `language` - Die Sprache für VLOOKUP-Evaluierung (z.B. Some("deutsch"))
/// * `e_income` - E-Spalte Einnahmen (für Check-Formel)
/// * `e_total` - E-Spalte Total (für Check-Formel)
/// * `f_income` - F-Spalte Einnahmen (für Check-Formel)
/// * `f_total` - F-Spalte Total (für Check-Formel)
/// * `bank` - Bank-Wert (für OK-Formel)
/// * `kasse` - Kasse-Wert (für OK-Formel)
/// * `sonstiges` - Sonstiges-Wert (für OK-Formel)
///
/// # Returns
/// Das berechnete Footer-Layout
#[allow(clippy::too_many_arguments)]
pub fn write_footer(
    ws: &mut Worksheet,
    fmt: &FormatMatrix,
    total_row: u32,
    income_row: u32,
    language: Option<&str>,
    e_income: Option<f64>,
    e_total: Option<f64>,
    f_income: Option<f64>,
    f_total: Option<f64>,
    bank: Option<f64>,
    kasse: Option<f64>,
    sonstiges: Option<f64>,
) -> Result<FooterLayout, XlsxError> {
    let layout = FooterLayout::compute(total_row);
    let s = layout.start_row;

    // =========================================================================
    // ZEILE 0 (s): "Kontrolle" / "Saldo für den Berichtszeitraum"
    // =========================================================================

    write_blank(ws, fmt, s, 1)?;
    write_blank(ws, fmt, s, 2)?;
    write_blank(ws, fmt, s, 3)?;

    // E(s):E(s+1) merged mit VLOOKUP 44 ("Saldo für den Berichtszeitraum")
    write_merged_vlookup_formula(ws, fmt, s, 4, s + 1, 4, 44, language)?;

    // =========================================================================
    // ZEILE 1 (s+1): "ABSCHLUSS"
    // =========================================================================

    // B:D(s+1) merged mit VLOOKUP 43 ("ABSCHLUSS")
    write_merged_vlookup_formula(ws, fmt, s + 1, 1, s + 1, 3, 43, language)?;
    // E ist bereits in merge von oben

    // =========================================================================
    // ZEILE 2 (s+2): Währungs-Referenz
    // =========================================================================

    write_blank(ws, fmt, s + 2, 1)?;
    write_blank(ws, fmt, s + 2, 2)?;
    write_blank(ws, fmt, s + 2, 3)?;

    // E: VLOOKUP(10) (Währung)
    write_vlookup_formula(ws, fmt, s + 2, 4, 10, language)?;

    // =========================================================================
    // ZEILE 3 (s+3): Leer
    // =========================================================================

    write_blank(ws, fmt, s + 3, 1)?;
    write_blank(ws, fmt, s + 3, 2)?;
    write_blank(ws, fmt, s + 3, 3)?;
    write_blank(ws, fmt, s + 3, 4)?;

    // =========================================================================
    // ZEILE 4 (s+4): Saldo-Differenz
    // =========================================================================

    // B: VLOOKUP 45 ("Saldo...")
    write_vlookup_formula(ws, fmt, s + 4, 1, 45, language)?;

    // C: blank
    write_blank(ws, fmt, s + 4, 2)?;

    // D: Check-Formel ✓
    let check_formula_str = format!(
        "=IF(ROUND(E{},2)=(ROUND(F{}-F{},2)),\"✓\",\"\")",
        s + 4 + 1,      // E saldo (1-indexed)
        income_row + 1, // F20 (1-indexed)
        total_row + 1   // F total (1-indexed)
    );
    let check_formula = match (e_income, e_total, f_income, f_total) {
        (Some(e_inc), Some(e_tot), Some(f_inc), Some(f_tot)) => {
            let e_diff = (e_inc - e_tot) * 100.0;
            let f_diff = (f_inc - f_tot) * 100.0;
            if e_diff.round() == f_diff.round() {
                Formula::new(&check_formula_str).set_result("✓")
            } else {
                Formula::new(&check_formula_str)
            }
        }
        _ => Formula::new(&check_formula_str),
    };
    write_formula_locked(ws, fmt, s + 4, 3, check_formula)?;

    // E: Differenz-Formel (E_income - E_total)
    let diff_formula_str = format!("=E{}-E{}", income_row + 1, total_row + 1);
    let diff_formula = match (e_income, e_total) {
        (Some(e_inc), Some(e_tot)) => {
            Formula::new(&diff_formula_str).set_result((e_inc - e_tot).to_string())
        }
        _ => Formula::new(&diff_formula_str),
    };
    write_formula_locked(ws, fmt, s + 4, 4, diff_formula)?;

    // =========================================================================
    // ZEILE 5 (s+5): Leer
    // =========================================================================

    write_blank(ws, fmt, s + 5, 1)?;
    write_blank(ws, fmt, s + 5, 2)?;
    write_blank(ws, fmt, s + 5, 3)?;
    write_blank(ws, fmt, s + 5, 4)?;

    // =========================================================================
    // ZEILE 6 (s+6): Saldenabstimmung
    // =========================================================================

    // B: VLOOKUP 46 ("Saldenabstimmung:")
    write_vlookup_formula(ws, fmt, s + 6, 1, 46, language)?;

    // C, D: blank
    write_blank(ws, fmt, s + 6, 2)?;
    write_blank(ws, fmt, s + 6, 3)?;

    // E: OK-Check
    let ok_formula_str = format!(
        "=IF(E{}=SUM(E{}:E{}),\"OK\",\"\")",
        s + 4 + 1, // saldo row (1-indexed)
        s + 7 + 1, // bank row (1-indexed)
        s + 9 + 1  // sonstiges row (1-indexed)
    );
    let ok_formula = match (e_income, e_total, bank, kasse, sonstiges) {
        (Some(e_inc), Some(e_tot), Some(b), Some(k), Some(so)) => {
            let e_saldo = e_inc - e_tot;
            let sum_inputs = b + k + so;
            if (e_saldo * 100.0).round() == (sum_inputs * 100.0).round() {
                Formula::new(&ok_formula_str).set_result("OK")
            } else {
                Formula::new(&ok_formula_str)
            }
        }
        _ => Formula::new(&ok_formula_str),
    };
    write_formula_locked(ws, fmt, s + 6, 4, ok_formula)?;

    // =========================================================================
    // ZEILEN 7-8 (s+7 bis s+8): Bank, Kasse
    // =========================================================================

    let vlookup_indices = [47, 48]; // Bank, Kasse
    for (i, vlookup_idx) in vlookup_indices.iter().enumerate() {
        let row = s + 7 + i as u32;

        write_vlookup_formula(ws, fmt, row, 1, *vlookup_idx, language)?;
        write_blank(ws, fmt, row, 2)?;
        write_blank(ws, fmt, row, 3)?;
        write_blank(ws, fmt, row, 4)?;
    }

    // =========================================================================
    // ZEILE 9 (s+9): Sonstiges - letzte Zeile mit bottom border
    // =========================================================================

    write_vlookup_formula(ws, fmt, s + 9, 1, 49, language)?;
    write_blank(ws, fmt, s + 9, 2)?;
    write_blank(ws, fmt, s + 9, 3)?;
    write_blank(ws, fmt, s + 9, 4)?;

    // =========================================================================
    // ZEILE 13 (s+13): Bestätigung 1
    // =========================================================================

    write_vlookup_formula(ws, fmt, s + 13, 1, 50, language)?;

    // =========================================================================
    // ZEILE 14 (s+14): Bestätigung 2
    // =========================================================================

    write_vlookup_formula(ws, fmt, s + 14, 1, 54, language)?;

    // =========================================================================
    // ZEILE 19 (s+19): Unterschriften
    // =========================================================================

    write_vlookup_formula(ws, fmt, s + 19, 1, 51, language)?;
    write_blank(ws, fmt, s + 19, 2)?;
    write_vlookup_formula(ws, fmt, s + 19, 3, 52, language)?;

    for col in 4..=6 {
        write_blank(ws, fmt, s + 19, col)?;
    }

    // =========================================================================
    // ZEILE 20 (s+20): Funktion
    // =========================================================================

    write_vlookup_formula(ws, fmt, s + 20, 3, 53, language)?;

    Ok(layout)
}

/// Schreibt die Footer Input-Werte (Bank, Kasse, Sonstiges)
pub fn write_footer_values(
    ws: &mut Worksheet,
    layout: &FooterLayout,
    fmt: &FormatMatrix,
    bank: Option<f64>,
    kasse: Option<f64>,
    sonstiges: Option<f64>,
) -> Result<(), XlsxError> {
    let s = layout.start_row;

    // Bank (E, Zeile 7)
    if let Some(value) = bank {
        if let Some(format) = fmt.get(s + 7, 4) {
            ws.write_number_with_format(s + 7, 4, value, format)?;
        }
    }

    // Kasse (E, Zeile 8)
    if let Some(value) = kasse {
        if let Some(format) = fmt.get(s + 8, 4) {
            ws.write_number_with_format(s + 8, 4, value, format)?;
        }
    }

    // Sonstiges (E, Zeile 9)
    if let Some(value) = sonstiges {
        if let Some(format) = fmt.get(s + 9, 4) {
            ws.write_number_with_format(s + 9, 4, value, format)?;
        }
    }

    Ok(())
}

// ============================================================================
// Helper-Funktionen
// ============================================================================

/// Schreibt einen Blank mit Format aus FormatMatrix
fn write_blank(
    ws: &mut Worksheet,
    fmt: &FormatMatrix,
    row: u32,
    col: u16,
) -> Result<(), XlsxError> {
    if let Some(format) = fmt.get(row, col) {
        ws.write_blank(row, col, format)?;
    }
    Ok(())
}

/// Schreibt eine VLOOKUP-Formel mit gecachtem Text-Ergebnis (locked)
fn write_vlookup_formula(
    ws: &mut Worksheet,
    fmt: &FormatMatrix,
    row: u32,
    col: u16,
    index: usize,
    language: Option<&str>,
) -> Result<(), XlsxError> {
    let formula_str = format!(
        r#"=IF($E$2="","",VLOOKUP($E$2,Sprachversionen!$B:$BN,{},FALSE))"#,
        index
    );

    let formula = if let Some(text) = lookup_text_string(language, index) {
        Formula::new(&formula_str).set_result(text)
    } else {
        Formula::new(&formula_str)
    };

    if let Some(format) = fmt.get_locked(row, col) {
        ws.write_formula_with_format(row, col, formula, &format)?;
    } else {
        ws.write_formula(row, col, formula)?;
    }
    Ok(())
}

/// Schreibt eine Formel mit locked Format aus FormatMatrix
fn write_formula_locked(
    ws: &mut Worksheet,
    fmt: &FormatMatrix,
    row: u32,
    col: u16,
    formula: Formula,
) -> Result<(), XlsxError> {
    if let Some(format) = fmt.get_locked(row, col) {
        ws.write_formula_with_format(row, col, formula, &format)?;
    } else {
        ws.write_formula(row, col, formula)?;
    }
    Ok(())
}

/// Schreibt eine VLOOKUP-Formel in einem gemergten Bereich mit gecachtem Text-Ergebnis (locked)
#[allow(clippy::too_many_arguments)]
fn write_merged_vlookup_formula(
    ws: &mut Worksheet,
    fmt: &FormatMatrix,
    row_start: u32,
    col_start: u16,
    row_end: u32,
    col_end: u16,
    index: usize,
    language: Option<&str>,
) -> Result<(), XlsxError> {
    let formula_str = format!(
        r#"=IF($E$2="","",VLOOKUP($E$2,Sprachversionen!$B:$BN,{},FALSE))"#,
        index
    );

    let formula = if let Some(text) = lookup_text_string(language, index) {
        Formula::new(&formula_str).set_result(text)
    } else {
        Formula::new(&formula_str)
    };

    if let Some(format) = fmt.get_locked(row_start, col_start) {
        ws.merge_range(row_start, col_start, row_end, col_end, "", &format)?;
        ws.write_formula_with_format(row_start, col_start, formula, &format)?;
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_footer_layout_compute() {
        // Total in Zeile 100 (0-indexed)
        let layout = FooterLayout::compute(100);

        assert_eq!(layout.start_row, 103); // 100 + 3
        assert_eq!(layout.saldo_row, 107); // 103 + 4
        assert_eq!(layout.input_rows[0], 110); // Bank: 103 + 7
        assert_eq!(layout.input_rows[1], 111); // Kasse: 103 + 8
        assert_eq!(layout.input_rows[2], 112); // Sonstiges: 103 + 9
        assert_eq!(layout.end_row, 123); // 103 + 20
    }
}
