//! Gemeinsame Hilfs-Funktionen für Section-Writer
//!
//! Diese Funktionen werden von mehreren Section-Writern verwendet:
//! - `write_blank`: Schreibt eine Blank-Zelle mit Format aus FormatMatrix
//! - `write_vlookup_formula`: Schreibt eine VLOOKUP-Formel mit gecachtem Text
//! - `write_formula_locked`: Schreibt eine beliebige Formel mit locked Format
//! - `write_merged_vlookup_formula`: Schreibt eine VLOOKUP-Formel in einem Merge-Bereich

use crate::report::core::lookup_text_string;
use crate::report::format::FormatMatrix;
use rust_xlsxwriter::{Formula, Worksheet, XlsxError};

/// Schreibt eine Blank-Zelle mit Format aus FormatMatrix.
/// Wenn kein Format für die Zelle existiert, wird nichts geschrieben.
pub fn write_blank(
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

/// Schreibt eine VLOOKUP-Formel mit gecachtem Text-Ergebnis (locked).
///
/// Die Formel lautet:
/// `=IF($E$2="","",VLOOKUP($E$2,Sprachversionen!$B:$BN,{index},FALSE))`
///
/// Das gecachte Text-Ergebnis wird via `lookup_text_string` vorberechnet,
/// sodass Excel den Text sofort anzeigt.
pub fn write_vlookup_formula(
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

/// Schreibt eine beliebige Formel mit locked Format aus FormatMatrix.
///
/// Verwendet `get_locked` — die Zelle wird als nicht editierbar markiert
/// wenn Sheet-Schutz aktiv ist.
pub fn write_formula_locked(
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

/// Schreibt eine VLOOKUP-Formel in einem gemergten Zellbereich (locked).
///
/// Merged den Bereich `(row_start, col_start):(row_end, col_end)` und schreibt
/// eine VLOOKUP-Formel mit gecachtem Text-Ergebnis in die erste Zelle.
#[allow(clippy::too_many_arguments)]
pub fn write_merged_vlookup_formula(
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
