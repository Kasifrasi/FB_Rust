//! Table Section - Neues zellbasiertes System
//!
//! Einnahmen-Tabelle (Rows 10-19, Excel 11-20)
//!
//! Struktur:
//! - Header (Rows 10-13): Vertikale Merges D11:D14, E11:E14, etc.
//! - Body (Rows 14-18): Horizontale Merges B15:C15, B16:C16, etc.
//! - Summary (Row 19): Horizontale Merge B20:C20

use crate::report::format::{BorderManager, BorderSpec, FillColors, MergedCellRegistry};
use rust_xlsxwriter::{Format, FormatAlign, Worksheet, XlsxError};

/// Schreibt Table Section mit neuem zellbasiertem System
///
/// # Workflow
///
/// 1. Merged Cells registrieren
/// 2. Borders konfigurieren
/// 3. Zellen schreiben mit gemerged Borders
///
/// # Arguments
///
/// * `ws` - Worksheet
/// * `fills` - Fill Colors
pub fn write_table_new(ws: &mut Worksheet, fills: &FillColors) -> Result<(), XlsxError> {
    // =========================================================================
    // PHASE 1: MERGED CELLS REGISTRIEREN (ZUERST!)
    // =========================================================================

    let mut merged_cells = MergedCellRegistry::new();

    // === HEADER MERGES (Rows 10-13) ===

    // B11:C11, B12:C12, B13:C13, B14:C14 (horizontale Merges für Labels)
    merged_cells.register_merge(10, 1, 10, 2); // B11:C11
    merged_cells.register_merge(11, 1, 11, 2); // B12:C12
    merged_cells.register_merge(12, 1, 12, 2); // B13:C13
    merged_cells.register_merge(13, 1, 13, 2); // B14:C14

    // D11:D14, E11:E14, F11:F14, G11:G14, H11:H14 (vertikale Merges für Spaltenüberschriften)
    merged_cells.register_merge(10, 3, 13, 3); // D11:D14
    merged_cells.register_merge(10, 4, 13, 4); // E11:E14
    merged_cells.register_merge(10, 5, 13, 5); // F11:F14
    merged_cells.register_merge(10, 6, 13, 6); // G11:G14
    merged_cells.register_merge(10, 7, 13, 7); // H11:H14

    // === BODY MERGES (Rows 14-18) ===

    for row in 14..=18 {
        merged_cells.register_merge(row, 1, row, 2); // B:C merge pro Zeile
    }

    // === SUMMARY MERGE (Row 19) ===

    merged_cells.register_merge(19, 1, 19, 2); // B20:C20

    // =========================================================================
    // PHASE 2: BORDERS KONFIGURIEREN (VOR dem Schreiben!)
    // =========================================================================

    let mut borders = BorderManager::new();
    borders.set_merged_cells(merged_cells.clone());

    // Box um gesamte Tabelle (B11:H20)
    borders.add_range(
        10,
        1,
        19,
        7,
        BorderSpec {
            top: Some(rust_xlsxwriter::FormatBorder::Medium),
            bottom: Some(rust_xlsxwriter::FormatBorder::Medium),
            left: Some(rust_xlsxwriter::FormatBorder::Medium),
            right: Some(rust_xlsxwriter::FormatBorder::Medium),
        },
    );

    // Horizontale Trennung zwischen Header und Body (Row 13/14)
    borders.add_range(
        13,
        1,
        13,
        7,
        BorderSpec {
            top: None,
            bottom: Some(rust_xlsxwriter::FormatBorder::Thin),
            left: None,
            right: None,
        },
    );

    // Horizontale Trennung zwischen Body und Summary (Row 18/19)
    borders.add_range(
        18,
        1,
        18,
        7,
        BorderSpec {
            top: None,
            bottom: Some(rust_xlsxwriter::FormatBorder::Medium),
            left: None,
            right: None,
        },
    );

    // Vertikale Trennungen (thin)
    for col in 2..7 {
        borders.add_range(
            10,
            col,
            19,
            col,
            BorderSpec {
                top: None,
                bottom: None,
                left: Some(rust_xlsxwriter::FormatBorder::Thin),
                right: None,
            },
        );
    }

    // =========================================================================
    // PHASE 3: ZELLEN SCHREIBEN (Zellbasiert, mit Border-Merge)
    // =========================================================================

    // --- HEADER (Rows 10-13, Excel 11-14) ---

    // B11:C11 (merged): Leer
    ws.merge_range(
        10,
        1,
        10,
        2,
        "",
        &Format::new()
            .set_font_name("Arial")
            .set_font_size(10.0)
            .set_align(FormatAlign::Center),
    )?;

    // D11:D14 (merged): Spaltenüberschrift (VLOOKUP würde hier stehen)
    ws.merge_range(
        10,
        3,
        13,
        3,
        "",
        &Format::new()
            .set_font_name("Arial")
            .set_font_size(10.0)
            .set_bold()
            .set_align(FormatAlign::Center)
            .set_text_wrap(),
    )?;

    // E11:E14 (merged): Spaltenüberschrift
    ws.merge_range(
        10,
        4,
        13,
        4,
        "",
        &Format::new()
            .set_font_name("Arial")
            .set_font_size(10.0)
            .set_bold()
            .set_align(FormatAlign::Center)
            .set_text_wrap(),
    )?;

    // F11:F14 (merged): Spaltenüberschrift
    ws.merge_range(
        10,
        5,
        13,
        5,
        "",
        &Format::new()
            .set_font_name("Arial")
            .set_font_size(10.0)
            .set_bold()
            .set_align(FormatAlign::Center)
            .set_text_wrap(),
    )?;

    // G11:G14 (merged): Spaltenüberschrift (Prozent)
    ws.merge_range(
        10,
        6,
        13,
        6,
        "",
        &Format::new()
            .set_font_name("Arial")
            .set_font_size(10.0)
            .set_bold()
            .set_align(FormatAlign::Center)
            .set_text_wrap(),
    )?;

    // H11:H14 (merged): Spaltenüberschrift
    ws.merge_range(
        10,
        7,
        13,
        7,
        "",
        &Format::new()
            .set_font_name("Arial")
            .set_font_size(10.0)
            .set_bold()
            .set_align(FormatAlign::Center)
            .set_text_wrap(),
    )?;

    // B12:C12 (merged): Zeilen-Label (VLOOKUP würde hier stehen)
    ws.merge_range(
        11,
        1,
        11,
        2,
        "",
        &Format::new()
            .set_font_name("Arial")
            .set_font_size(10.0)
            .set_align(FormatAlign::Center),
    )?;

    // B13:C13 (merged): Zeilen-Label
    ws.merge_range(
        12,
        1,
        12,
        2,
        "",
        &Format::new()
            .set_font_name("Arial")
            .set_font_size(10.0)
            .set_align(FormatAlign::Center),
    )?;

    // B14:C14 (merged): Leer
    ws.merge_range(
        13,
        1,
        13,
        2,
        "",
        &Format::new()
            .set_font_name("Arial")
            .set_font_size(10.0)
            .set_align(FormatAlign::Center),
    )?;

    // --- BODY (Rows 14-18, Excel 15-19) ---

    for row in 14..=18 {
        // B:C (merged): Label (VLOOKUP würde hier stehen)
        ws.merge_range(
            row,
            1,
            row,
            2,
            "",
            &Format::new()
                .set_font_name("Arial")
                .set_font_size(10.0)
                .set_align(FormatAlign::Left),
        )?;

        // D: Number (Input, gelb, unlocked)
        ws.write_number_with_format(
            row,
            3,
            0.0,
            &Format::new()
                .set_font_name("Arial")
                .set_font_size(10.0)
                .set_background_color(fills.input) // ← Fill via Variable!
                .set_align(FormatAlign::Right)
                .set_unlocked()
                .set_num_format("#,##0.00"),
        )?;

        // E: Number (Input, gelb, unlocked)
        ws.write_number_with_format(
            row,
            4,
            0.0,
            &Format::new()
                .set_font_name("Arial")
                .set_font_size(10.0)
                .set_background_color(fills.input) // ← Fill via Variable!
                .set_align(FormatAlign::Right)
                .set_unlocked()
                .set_num_format("#,##0.00"),
        )?;

        // F: Calculated (Formel, grau, locked)
        ws.write_blank(
            row,
            5,
            &Format::new()
                .set_font_name("Arial")
                .set_font_size(10.0)
                .set_background_color(fills.value) // ← Fill via Variable!
                .set_align(FormatAlign::Right)
                .set_num_format("#,##0.00"),
        )?;

        // G: Percentage (Formel, locked)
        ws.write_blank(
            row,
            6,
            &Format::new()
                .set_font_name("Arial")
                .set_font_size(10.0)
                .set_align(FormatAlign::Center)
                .set_num_format("0.00%"),
        )?;

        // H: Right column (Blank)
        ws.write_blank(
            row,
            7,
            &Format::new()
                .set_font_name("Arial")
                .set_font_size(10.0)
                .set_align(FormatAlign::Right),
        )?;
    }

    // --- SUMMARY ROW (Row 19, Excel 20) ---

    // B20:C20 (merged): "Summe" Label
    ws.merge_range(
        19,
        1,
        19,
        2,
        "",
        &Format::new()
            .set_font_name("Arial")
            .set_font_size(10.0)
            .set_bold()
            .set_align(FormatAlign::Center),
    )?;

    // D20: Sum (Formel, grau, locked)
    ws.write_blank(
        19,
        3,
        &Format::new()
            .set_font_name("Arial")
            .set_font_size(10.0)
            .set_background_color(fills.summary) // ← Fill via Variable!
            .set_bold()
            .set_align(FormatAlign::Right)
            .set_num_format("#,##0.00"),
    )?;

    // E20: Sum (Formel, grau, locked)
    ws.write_blank(
        19,
        4,
        &Format::new()
            .set_font_name("Arial")
            .set_font_size(10.0)
            .set_background_color(fills.summary) // ← Fill via Variable!
            .set_bold()
            .set_align(FormatAlign::Right)
            .set_num_format("#,##0.00"),
    )?;

    // F20: Sum (Formel, grau, locked)
    ws.write_blank(
        19,
        5,
        &Format::new()
            .set_font_name("Arial")
            .set_font_size(10.0)
            .set_background_color(fills.summary) // ← Fill via Variable!
            .set_bold()
            .set_align(FormatAlign::Right)
            .set_num_format("#,##0.00"),
    )?;

    // G20: Percentage (Formel, locked)
    ws.write_blank(
        19,
        6,
        &Format::new()
            .set_font_name("Arial")
            .set_font_size(10.0)
            .set_bold()
            .set_align(FormatAlign::Center)
            .set_num_format("0.00%"),
    )?;

    // H20: Right column (Blank)
    ws.write_blank(
        19,
        7,
        &Format::new()
            .set_font_name("Arial")
            .set_font_size(10.0)
            .set_align(FormatAlign::Right),
    )?;

    Ok(())
}

/// Merged ein Format mit Borders falls für die Zelle vorhanden
///
/// Helper-Funktion für das Baukastenprinzip: Format + Border = Finales Format
fn with_borders(format: Format, borders: &BorderManager, row: u32, col: u16) -> Format {
    if let Ok(Some(border_fmt)) = borders.get_border_for_cell(row, col) {
        format.merge(&border_fmt)
    } else {
        format
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use rust_xlsxwriter::Workbook;

    #[test]
    fn test_table_new_compiles() {
        let mut wb = Workbook::new();
        let ws = wb.add_worksheet();
        let fills = FillColors::new();

        let result = write_table_new(ws, &fills);
        assert!(result.is_ok());
    }

    #[test]
    fn test_merged_cells_count() {
        let mut merged_cells = MergedCellRegistry::new();

        // Header: 4 horizontal + 5 vertical = 9
        merged_cells.register_merge(10, 1, 10, 2);
        merged_cells.register_merge(11, 1, 11, 2);
        merged_cells.register_merge(12, 1, 12, 2);
        merged_cells.register_merge(13, 1, 13, 2);

        merged_cells.register_merge(10, 3, 13, 3);
        merged_cells.register_merge(10, 4, 13, 4);
        merged_cells.register_merge(10, 5, 13, 5);
        merged_cells.register_merge(10, 6, 13, 6);
        merged_cells.register_merge(10, 7, 13, 7);

        // Body: 5 rows
        for row in 14..=18 {
            merged_cells.register_merge(row, 1, row, 2);
        }

        // Summary: 1
        merged_cells.register_merge(19, 1, 19, 2);

        assert_eq!(merged_cells.count(), 15);
    }
}
