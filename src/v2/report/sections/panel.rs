//! Right Panel Section (Zeilen 10-30, Spalten J-V)
//!
//! Enthält zwei Panels:
//! - Linke Seite: J-O (Cols 9-14)
//! - Rechte Seite: Q-V (Cols 16-21)
//!
//! Struktur pro Panel:
//! - Header (Row 10-12): J11:K11 merged, L13-O13 Formeln
//! - Body (Row 13-30): 18 Zeilen mit Index, Text, Date, Numbers, Calc
//!
//! **Hinweis:** Header-Formeln (J11, Q11, L13-O13, S13-V13) werden von
//! `write_cells_from_registry()` geschrieben. Body-Rows werden hier geschrieben.

use crate::v2::lang::data::TEXT_MATRIX;
use crate::v2::report::formats::FormatMatrix;
use crate::v2::report::layout::MergeRange;
use crate::v2::report::values::ReportValues;
use rust_xlsxwriter::{Formula, Worksheet, XlsxError};

/// Merge-Bereiche für Panel Header
pub const HEADER_MERGES: &[(MergeRange, u32, u16)] = &[
    (MergeRange::new(10, 9, 10, 10), 10, 9),   // J11:K11
    (MergeRange::new(10, 16, 10, 17), 10, 16), // Q11:R11
];

/// Schreibt die Panel Section (Layout, Merges, Blanks, Body-Rows)
///
/// **Hinweis:** Header-Formeln (J11, Q11, L13-O13, S13-V13) werden von
/// `write_cells_from_registry()` geschrieben. Body-Rows (K/R Spalten mit
/// VLOOKUP und O/V Calc-Formeln) werden hier direkt geschrieben, da sie
/// in jeder Zeile wiederholt werden.
pub fn write_panel_section(
    ws: &mut Worksheet,
    fmt: &FormatMatrix,
    values: &ReportValues,
) -> Result<(), XlsxError> {
    // Text-Lookup für K/R Spalten auswerten
    let text_result = evaluate_text_lookup_23(values);
    let f_k = r#"=IF($E$2="","",VLOOKUP($E$2,Sprachversionen!$B:$BN,23,FALSE))"#;

    // Header Merges
    write_header_merges(ws, fmt)?;

    // Header Blanks (Row 11-12)
    write_header_blanks(ws, fmt)?;

    // Header Formeln (J11, Q11, L13-O13, S13-V13) werden von
    // write_cells_from_registry() geschrieben!

    // Body Rows (Row 13-30) - K/R VLOOKUP und O/V Calc-Formeln
    // Diese werden hier geschrieben, da sie in jeder Zeile wiederholt werden
    // und die O/V Calc-Formeln bereits in der Registry registriert sind.
    write_body_rows(ws, fmt, f_k, &text_result)?;

    Ok(())
}

fn write_header_merges(ws: &mut Worksheet, fmt: &FormatMatrix) -> Result<(), XlsxError> {
    for (range, fmt_row, fmt_col) in HEADER_MERGES {
        if let Some(format) = fmt.get(*fmt_row, *fmt_col) {
            ws.merge_range(
                range.first_row,
                range.first_col,
                range.last_row,
                range.last_col,
                "",
                format,
            )?;
        }
    }
    Ok(())
}

fn write_header_blanks(ws: &mut Worksheet, fmt: &FormatMatrix) -> Result<(), XlsxError> {
    // Row 11 (Index 10): L11-O11 und S11-V11 sind blanks
    for col in [11u16, 12, 13, 14] {
        if let Some(format) = fmt.get(10, col) {
            ws.write_blank(10, col, format)?;
        }
    }
    for col in [18u16, 19, 20, 21] {
        if let Some(format) = fmt.get(10, col) {
            ws.write_blank(10, col, format)?;
        }
    }

    // Row 12 (Index 11): Alle J12-O12 und Q12-V12 sind blanks
    for col in [9u16, 10, 11, 12, 13, 14] {
        if let Some(format) = fmt.get(11, col) {
            ws.write_blank(11, col, format)?;
        }
    }
    for col in [16u16, 17, 18, 19, 20, 21] {
        if let Some(format) = fmt.get(11, col) {
            ws.write_blank(11, col, format)?;
        }
    }

    // Row 13 (Index 12): Nur J13, K13, Q13, R13 sind blanks
    // L13-O13 und S13-V13 haben Formeln (werden in writer.rs geschrieben)
    for col in [9u16, 10] {
        if let Some(format) = fmt.get(12, col) {
            ws.write_blank(12, col, format)?;
        }
    }
    for col in [16u16, 17] {
        if let Some(format) = fmt.get(12, col) {
            ws.write_blank(12, col, format)?;
        }
    }

    Ok(())
}

fn write_body_rows(
    ws: &mut Worksheet,
    fmt: &FormatMatrix,
    f_k: &str,
    text_result: &str,
) -> Result<(), XlsxError> {
    // 18 Zeilen: Row 13-30 (Index 13-30)
    // Links: Index 1-18, Rechts: Index 19-36
    for i in 0..18u32 {
        let row = 13 + i;
        let left_num = 1 + i;
        let right_num = 19 + i;
        write_panel_row(ws, fmt, row, left_num, right_num, f_k, text_result)?;
    }
    Ok(())
}

fn write_panel_row(
    ws: &mut Worksheet,
    fmt: &FormatMatrix,
    row: u32,
    left_num: u32,
    right_num: u32,
    f_k: &str,
    text_result: &str,
) -> Result<(), XlsxError> {
    // === LINKE SEITE (J-O) ===

    // J: Index
    if let Some(format) = fmt.get(row, 9) {
        ws.write_string_with_format(row, 9, &format!("{}. ", left_num), format)?;
    }

    // K: Text (Formel mit ausgewertetem Ergebnis)
    let formula_k = Formula::new(f_k).set_result(text_result);
    if let Some(format) = fmt.get_locked(row, 10) {
        ws.write_formula_with_format(row, 10, formula_k, &format)?;
    }

    // L: Date (blank)
    if let Some(format) = fmt.get(row, 11) {
        ws.write_blank(row, 11, format)?;
    }

    // M: Number (blank)
    if let Some(format) = fmt.get(row, 12) {
        ws.write_blank(row, 12, format)?;
    }

    // N: Number (blank)
    if let Some(format) = fmt.get(row, 13) {
        ws.write_blank(row, 13, format)?;
    }

    // O: Calc (Formel)
    let f_o = format!("=IF(M{}=\"\",\"\",N{}/M{})", row + 1, row + 1, row + 1);
    let formula_o = Formula::new(&f_o).set_result("");
    if let Some(format) = fmt.get_locked(row, 14) {
        ws.write_formula_with_format(row, 14, formula_o, &format)?;
    }

    // === RECHTE SEITE (Q-V) ===

    // Q: Index
    if let Some(format) = fmt.get(row, 16) {
        ws.write_string_with_format(row, 16, &format!("{}. ", right_num), format)?;
    }

    // R: Text (Formel mit ausgewertetem Ergebnis)
    let formula_r = Formula::new(f_k).set_result(text_result);
    if let Some(format) = fmt.get_locked(row, 17) {
        ws.write_formula_with_format(row, 17, formula_r, &format)?;
    }

    // S: Date (blank)
    if let Some(format) = fmt.get(row, 18) {
        ws.write_blank(row, 18, format)?;
    }

    // T: Number (blank)
    if let Some(format) = fmt.get(row, 19) {
        ws.write_blank(row, 19, format)?;
    }

    // U: Number (blank)
    if let Some(format) = fmt.get(row, 20) {
        ws.write_blank(row, 20, format)?;
    }

    // V: Calc (Formel)
    let f_v = format!("=IF(T{}=\"\",\"\",U{}/T{})", row + 1, row + 1, row + 1);
    let formula_v = Formula::new(&f_v).set_result("");
    if let Some(format) = fmt.get_locked(row, 21) {
        ws.write_formula_with_format(row, 21, formula_v, &format)?;
    }

    Ok(())
}

/// Evaluiert Text-Lookup für Index 23 (Right Panel K/R Spalten)
fn evaluate_text_lookup_23(values: &ReportValues) -> String {
    let language = match values.language() {
        Some(lang) if !lang.is_empty() => lang,
        _ => return String::new(),
    };

    let lang_idx = TEXT_MATRIX
        .iter()
        .position(|row| !row.is_empty() && row[0].eq_ignore_ascii_case(language));

    match lang_idx {
        Some(idx) => TEXT_MATRIX
            .get(idx)
            .and_then(|row| row.get(22)) // Index 23 in Excel = Index 22 in 0-basiertem Array
            .map(|s| s.to_string())
            .unwrap_or_default(),
        None => String::new(),
    }
}
