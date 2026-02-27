//! Right Panel Section (Zeilen 10-30, Spalten J-V)
//!
//! Enthält zwei Panels:
//! - Linke Seite: J-O (Cols 9-14)
//! - Rechte Seite: Q-V (Cols 16-21)
//!
//! Struktur pro Panel:
//! - Header (Row 10-12): J11:K11 merged, L13-O13 Formeln
//! - Body (Row 13-30): 18 Zeilen mit Index, Blanks
//!
//! **Hinweis:** Alle Formeln werden von `write_cells_from_registry()` geschrieben:
//! - Header: J11, Q11, L13-O13, S13-V13
//! - Body: K14-K31 (VLOOKUP(23)), O14-O31 (Calc), R14-R31 (VLOOKUP(23)), V14-V31 (Calc)
//!   Dieser Writer schreibt nur Index-Strings (J/Q) und Blank-Zellen (L/M/N/S/T/U).

use super::super::layout::MergeRange;
use crate::report::format::FormatMatrix;
use rust_xlsxwriter::{Worksheet, XlsxError};

/// Merge-Bereiche für Panel Header
pub const HEADER_MERGES: &[(MergeRange, u32, u16)] = &[
    (MergeRange::new(10, 9, 10, 10), 10, 9),   // J11:K11
    (MergeRange::new(10, 16, 10, 17), 10, 16), // Q11:R11
];

/// Schreibt die Panel Section (Header-Merges, Blanks, Body Index-Strings)
///
/// **Hinweis:** Alle Formeln (K/O/R/V Body + Header) werden von
/// `write_cells_from_registry()` geschrieben.
pub fn write_panel_section(ws: &mut Worksheet, fmt: &FormatMatrix) -> Result<(), XlsxError> {
    // Header Merges
    write_header_merges(ws, fmt)?;

    // Header Blanks (Row 11-12)
    write_header_blanks(ws, fmt)?;

    // Body Rows (Row 13-30): nur Index-Strings und Blanks
    write_body_rows(ws, fmt)?;

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

    // Row 13 (Index 12): J13, K13, Q13, R13 sind blanks
    // L13-O13 und S13-V13 haben Formeln (werden von write_cells_from_registry geschrieben)
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

fn write_body_rows(ws: &mut Worksheet, fmt: &FormatMatrix) -> Result<(), XlsxError> {
    // 18 Zeilen: Row 13-30 (Index 13-30)
    // Links: Index 1-18, Rechts: Index 19-36
    for i in 0..18u32 {
        let row = 13 + i;
        let left_num = 1 + i;
        let right_num = 19 + i;
        write_panel_row(ws, fmt, row, left_num, right_num)?;
    }
    Ok(())
}

fn write_panel_row(
    ws: &mut Worksheet,
    fmt: &FormatMatrix,
    row: u32,
    left_num: u32,
    right_num: u32,
) -> Result<(), XlsxError> {
    // === LINKE SEITE (J-O) ===

    // J: Index
    if let Some(format) = fmt.get(row, 9) {
        ws.write_string_with_format(row, 9, format!("{}. ", left_num), format)?;
    }

    // K: VLOOKUP(23) — via Registry

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

    // O: Calc — via Registry

    // === RECHTE SEITE (Q-V) ===

    // Q: Index
    if let Some(format) = fmt.get(row, 16) {
        ws.write_string_with_format(row, 16, format!("{}. ", right_num), format)?;
    }

    // R: VLOOKUP(23) — via Registry

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

    // V: Calc — via Registry

    Ok(())
}
