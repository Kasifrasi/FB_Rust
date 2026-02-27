//! Report Footer Sektion - Nur Struktur
//!
//! Schreibt nur statische Strukturelemente (Merges, Blanks).
//! Alle Formeln (VLOOKUP-Labels, Check/Diff/OK) und API-Werte (Bank, Kasse, Sonstiges)
//! werden von `write_cells_from_registry()` geschrieben.
//!
//! Struktur (21 Zeilen ab start_row):
//! - Zeile 0: Blanks B-D, E:E+1 Merge
//! - Zeile 1: B:D Merge
//! - Zeile 2: Blanks B-D
//! - Zeile 3: Blanks B-E
//! - Zeile 4: Blank C
//! - Zeile 5: Blanks B-E
//! - Zeile 6: Blanks C-D
//! - Zeile 7-8: Blanks C-E
//! - Zeile 9: Blanks C-E
//! - Zeile 19: Blank C, Blanks E-G
//! - Zeile 13,14,19,20: Labels via Registry

use super::utils::write_blank;
use crate::report::body::FooterLayout;
use crate::report::format::FormatMatrix;
use rust_xlsxwriter::{Worksheet, XlsxError};

/// Schreibt die Footer-Struktur (nur Merges und Blanks)
///
/// Alle Formeln und API-Werte werden von `write_cells_from_registry()` geschrieben.
pub fn write_footer_structure(
    ws: &mut Worksheet,
    fmt: &FormatMatrix,
    layout: &FooterLayout,
) -> Result<(), XlsxError> {
    let s = layout.start_row;

    // =========================================================================
    // ZEILE 0 (s): Blanks + E:E+1 Merge
    // =========================================================================

    write_blank(ws, fmt, s, 1)?;
    write_blank(ws, fmt, s, 2)?;
    write_blank(ws, fmt, s, 3)?;

    // E(s):E(s+1) merged (VLOOKUP 44 wird von Registry geschrieben)
    if let Some(format) = fmt.get_locked(s, 4) {
        ws.merge_range(s, 4, s + 1, 4, "", &format)?;
    }

    // =========================================================================
    // ZEILE 1 (s+1): B:D Merge
    // =========================================================================

    // B:D(s+1) merged (VLOOKUP 43 wird von Registry geschrieben)
    if let Some(format) = fmt.get_locked(s + 1, 1) {
        ws.merge_range(s + 1, 1, s + 1, 3, "", &format)?;
    }

    // =========================================================================
    // ZEILE 2 (s+2): Blanks
    // =========================================================================

    write_blank(ws, fmt, s + 2, 1)?;
    write_blank(ws, fmt, s + 2, 2)?;
    write_blank(ws, fmt, s + 2, 3)?;

    // =========================================================================
    // ZEILE 3 (s+3): Blanks
    // =========================================================================

    write_blank(ws, fmt, s + 3, 1)?;
    write_blank(ws, fmt, s + 3, 2)?;
    write_blank(ws, fmt, s + 3, 3)?;
    write_blank(ws, fmt, s + 3, 4)?;

    // =========================================================================
    // ZEILE 4 (s+4): Blank C
    // =========================================================================

    // B: VLOOKUP 45 via Registry
    // C: blank
    write_blank(ws, fmt, s + 4, 2)?;
    // D: Check-Formel via Registry
    // E: Diff-Formel via Registry

    // =========================================================================
    // ZEILE 5 (s+5): Blanks
    // =========================================================================

    write_blank(ws, fmt, s + 5, 1)?;
    write_blank(ws, fmt, s + 5, 2)?;
    write_blank(ws, fmt, s + 5, 3)?;
    write_blank(ws, fmt, s + 5, 4)?;

    // =========================================================================
    // ZEILE 6 (s+6): Blanks C, D
    // =========================================================================

    // B: VLOOKUP 46 via Registry
    write_blank(ws, fmt, s + 6, 2)?;
    write_blank(ws, fmt, s + 6, 3)?;
    // E: OK-Check via Registry

    // =========================================================================
    // ZEILEN 7-8 (s+7, s+8): Bank, Kasse — Blanks C, D, E(input)
    // =========================================================================

    for i in 7..=8 {
        let row = s + i;
        // B: VLOOKUP via Registry
        write_blank(ws, fmt, row, 2)?;
        write_blank(ws, fmt, row, 3)?;
        // E: API-Wert via Registry (Bank/Kasse)
    }

    // =========================================================================
    // ZEILE 9 (s+9): Sonstiges — Blanks C, D, E(input)
    // =========================================================================

    // B: VLOOKUP 49 via Registry
    write_blank(ws, fmt, s + 9, 2)?;
    write_blank(ws, fmt, s + 9, 3)?;
    // E: API-Wert via Registry (Sonstiges)

    // =========================================================================
    // ZEILE 19 (s+19): Unterschriften — Blank C, Blanks E-G
    // =========================================================================

    // B: VLOOKUP 51 via Registry
    write_blank(ws, fmt, s + 19, 2)?;
    // D: VLOOKUP 52 via Registry

    for col in 4..=6 {
        write_blank(ws, fmt, s + 19, col)?;
    }

    // ZEILE 20 (s+20): D VLOOKUP 53 via Registry — keine Blanks nötig

    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_footer_layout_compute() {
        let layout = FooterLayout::compute(100);

        assert_eq!(layout.start_row, 103);
        assert_eq!(layout.saldo_row, 107);
        assert_eq!(layout.input_rows[0], 110);
        assert_eq!(layout.input_rows[1], 111);
        assert_eq!(layout.input_rows[2], 112);
        assert_eq!(layout.end_row, 123);
    }
}
