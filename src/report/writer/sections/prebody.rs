//! Pre-Body Section (Zeilen 20-25)
//!
//! Statischer Bereich zwischen Einnahme-Tabelle und dynamischem Body.
//! Enthält:
//! - Spacer Row 20-21
//! - Merge-Struktur (Rows 22-25) mit vertikalen Merges
//!
//! ## Registry-Integration
//!
//! Formeln werden von `write_cells_from_registry` geschrieben
//! (via `register_prebody_formulas` in `definitions.rs`).
//! Diese Funktion schreibt nur noch das Layout (Merges, Blanks).

use super::utils::write_blank;
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
pub fn write_prebody_section(ws: &mut Worksheet, fmt: &FormatMatrix) -> Result<(), XlsxError> {
    // Spacer Row 20 (0-basiert: 20 = Excel 21)
    ws.set_row_height(20, 13.5)?;

    // Merge-Struktur für Rows 22-25
    write_prebody_merges(ws, fmt)?;

    Ok(())
}

fn write_prebody_merges(ws: &mut Worksheet, fmt: &FormatMatrix) -> Result<(), XlsxError> {
    // D-H (Spalten 3-7): Vertikale Merges Zeile 22-25
    // Formeln werden von der Registry in die Top-Left-Zelle geschrieben
    for col in 3u16..=7u16 {
        if let Some(format) = fmt.get_locked(22, col) {
            ws.merge_range(22, col, 25, col, "", &format)?;
        }
    }

    // B23, C23: Blanks
    write_blank(ws, fmt, 22, 1)?;
    write_blank(ws, fmt, 22, 2)?;

    // B24:C24 merged (Row 23, 0-indexed) — Formel via Registry
    if let Some(format) = fmt.get_locked(23, 1) {
        ws.merge_range(23, 1, 23, 2, "", &format)?;
    }

    // B25:C25 merged (Row 24, 0-indexed) — Formel via Registry
    if let Some(format) = fmt.get_locked(24, 1) {
        ws.merge_range(24, 1, 24, 2, "", &format)?;
    }

    // B26:C26 merged blank mit thin bottom border (Row 25, kein Formel)
    if let Some(format) = fmt.get(25, 1) {
        ws.merge_range(25, 1, 25, 2, "", format)?;
    }

    Ok(())
}
