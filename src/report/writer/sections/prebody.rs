//! Pre-Body Section (Zeilen 20-25)
//!
//! Statischer Bereich zwischen Einnahme-Tabelle und dynamischem Body.
//! Enthält:
//! - Spacer Row 20-21
//! - Footer-Tabelle (Rows 22-25) mit vertikalen Merges
//!
//! ## Registry-Integration
//!
//! Die VLOOKUP-Formeln in dieser Section sind jetzt in der zentralen
//! CellRegistry registriert (siehe `definitions.rs::register_prebody_formulas`).
//! Die Funktion `write_prebody_section` schreibt nur noch das Layout (Merges, Blanks),
//! während die Formeln von `write_cells_from_registry` geschrieben werden.

use super::utils::{write_blank, write_merged_vlookup_formula};
use crate::report::format::FormatMatrix;
use rust_xlsxwriter::{Worksheet, XlsxError};

/// Schreibt die Pre-Body Section (Zeilen 20-25)
///
/// Formate werden aus der FormatMatrix gelesen — konsistent mit allen anderen
/// Section-Writern.
///
/// # Arguments
/// * `ws` - Das Worksheet
/// * `fmt` - FormatMatrix (befüllt durch `extend_format_matrix_with_prebody`)
/// * `language` - Die Sprache für VLOOKUP-Evaluierung (z.B. Some("deutsch"))
pub fn write_prebody_section(
    ws: &mut Worksheet,
    fmt: &FormatMatrix,
    language: Option<&str>,
) -> Result<(), XlsxError> {
    // Spacer Row 20 (0-basiert: 20 = Excel 21)
    ws.set_row_height(20, 13.5)?;

    // Footer-Tabelle (Rows 22-25, 0-basiert: 22-25)
    write_footer_table(ws, fmt, language)?;

    Ok(())
}

/// Schreibt die Footer-Tabelle (Zeilen 22-25)
fn write_footer_table(
    ws: &mut Worksheet,
    fmt: &FormatMatrix,
    language: Option<&str>,
) -> Result<(), XlsxError> {
    // === Row 22 (Excel 23): Spaltenüberschriften mit vertikalen Merges ===
    // Columns D-H (3-7) — jede Spalte geht von Zeile 22 bis Zeile 25
    for (col, idx) in [(3u16, 11usize), (4, 25), (5, 55), (6, 56), (7, 15)] {
        write_merged_vlookup_formula(ws, fmt, 22, col, 25, col, idx, language)?;
    }

    // B23, C23 - Blanks
    write_blank(ws, fmt, 22, 1)?;
    write_blank(ws, fmt, 22, 2)?;

    // === Row 23 (Excel 24): B24:C24 merged - VLOOKUP(24) ===
    write_merged_vlookup_formula(ws, fmt, 23, 1, 23, 2, 24, language)?;

    // === Row 24 (Excel 25): B25:C25 merged - VLOOKUP(10) (Währung) ===
    write_merged_vlookup_formula(ws, fmt, 24, 1, 24, 2, 10, language)?;

    // === Row 25 (Excel 26): B26:C26 merged blank (thin bottom border) ===
    if let Some(format) = fmt.get(25, 1) {
        ws.merge_range(25, 1, 25, 2, "", format)?;
    }

    Ok(())
}
