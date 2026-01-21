//! Pre-Body Section - Neues zellbasiertes System
//!
//! Finance Table (Zeilen 22-25, Excel 23-26) mit vielen Merged Cells.
//!
//! Perfektes Beispiel für:
//! - Vertikale Merges (D23:D26, E23:E26, F23:F26, G23:G26, H23:H26)
//! - Horizontale Merges (B24:C24, B25:C25, B26:C26)
//! - Borders mit Merge-Awareness

use crate::report::core::lookup_text_string;
use crate::report::format::{BorderManager, BorderSpec, FillColors, MergedCellRegistry};
use rust_xlsxwriter::{Format, FormatAlign, Formula, Worksheet, XlsxError};

/// Schreibt PreBody Section mit neuem zellbasiertem System
///
/// # Workflow
///
/// 1. Merged Cells registrieren (ZUERST!)
/// 2. Borders konfigurieren (VOR dem Schreiben!)
/// 3. Zellen schreiben mit gemerged Borders
///
/// # Arguments
///
/// * `ws` - Worksheet
/// * `language` - Sprache für VLOOKUP (optional)
/// * `fills` - Fill Colors
pub fn write_prebody_new(
    ws: &mut Worksheet,
    language: Option<&str>,
    _fills: &FillColors,
) -> Result<(), XlsxError> {
    // Spacer Row 20 (Excel 21)
    ws.set_row_height(20, 13.5)?;

    // =========================================================================
    // PHASE 1: MERGED CELLS REGISTRIEREN (ZUERST!)
    // =========================================================================

    let mut merged_cells = MergedCellRegistry::new();

    // Vertikale Merges (Spaltenüberschriften, Rows 22-25 = Excel 23-26)
    merged_cells.register_merge(22, 3, 25, 3); // D23:D26
    merged_cells.register_merge(22, 4, 25, 4); // E23:E26
    merged_cells.register_merge(22, 5, 25, 5); // F23:F26
    merged_cells.register_merge(22, 6, 25, 6); // G23:G26
    merged_cells.register_merge(22, 7, 25, 7); // H23:H26

    // Horizontale Merges (Zeilen-Labels)
    merged_cells.register_merge(23, 1, 23, 2); // B24:C24
    merged_cells.register_merge(24, 1, 24, 2); // B25:C25
    merged_cells.register_merge(25, 1, 25, 2); // B26:C26

    // =========================================================================
    // PHASE 2: BORDERS KONFIGURIEREN (VOR dem Schreiben!)
    // =========================================================================

    let mut borders = BorderManager::new();
    borders.set_merged_cells(merged_cells.clone());

    // Box um gesamte Tabelle (B23:H26)
    borders.add_range(
        22,
        1,
        25,
        7,
        BorderSpec {
            top: Some(rust_xlsxwriter::FormatBorder::Medium),
            bottom: Some(rust_xlsxwriter::FormatBorder::Thin),
            left: Some(rust_xlsxwriter::FormatBorder::Medium),
            right: Some(rust_xlsxwriter::FormatBorder::Medium),
        },
    );

    // Vertikale Trennung zwischen B-C und D-H (thin)
    borders.add_range(
        22,
        2,
        25,
        2,
        BorderSpec {
            top: None,
            bottom: None,
            left: None,
            right: Some(rust_xlsxwriter::FormatBorder::Thin),
        },
    );

    // Vertikale Trennungen zwischen Spalten D-H (thin)
    for col in 3..7 {
        borders.add_range(
            22,
            col,
            25,
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

    // --- ROW 22 (Excel 23): Spaltenüberschriften ---

    // B23, C23: Leer (ohne Merge)
    ws.write_blank(
        22,
        1,
        &Format::new()
            .set_font_name("Arial")
            .set_font_size(10.0)
            .set_align(FormatAlign::Center),
    )?;

    ws.write_blank(
        22,
        2,
        &Format::new()
            .set_font_name("Arial")
            .set_font_size(10.0)
            .set_align(FormatAlign::Center),
    )?;

    // D23:D26 (merged): VLOOKUP(11)
    let fmt_d23 = with_borders(
        Format::new()
            .set_font_name("Arial")
            .set_font_size(10.0)
            .set_align(FormatAlign::Center)
            .set_text_wrap(),
        &borders,
        22,
        3,
    );

    let formula_11 = make_vlookup_formula(11, language);
    ws.merge_range(22, 3, 25, 3, "", &fmt_d23)?;
    ws.write_formula_with_format(22, 3, formula_11, &fmt_d23)?;

    // E23:E26 (merged): VLOOKUP(25) "Ausgaben" - BOLD
    let fmt_e23 = Format::new()
        .set_font_name("Arial")
        .set_font_size(10.0)
        .set_bold() // ← BOLD für "Ausgaben"
        .set_align(FormatAlign::Center)
        .set_text_wrap();

    let formula_25 = make_vlookup_formula(25, language);
    ws.merge_range(22, 4, 25, 4, "", &fmt_e23)?;
    ws.write_formula_with_format(22, 4, formula_25, &fmt_e23)?;

    // F23:F26 (merged): VLOOKUP(55)
    let fmt_f23 = Format::new()
        .set_font_name("Arial")
        .set_font_size(10.0)
        .set_align(FormatAlign::Center)
        .set_text_wrap();

    let formula_55 = make_vlookup_formula(55, language);
    ws.merge_range(22, 5, 25, 5, "", &fmt_f23)?;
    ws.write_formula_with_format(22, 5, formula_55, &fmt_f23)?;

    // G23:G26 (merged): VLOOKUP(56)
    let fmt_g23 = Format::new()
        .set_font_name("Arial")
        .set_font_size(10.0)
        .set_align(FormatAlign::Center)
        .set_text_wrap();

    let formula_56 = make_vlookup_formula(56, language);
    ws.merge_range(22, 6, 25, 6, "", &fmt_g23)?;
    ws.write_formula_with_format(22, 6, formula_56, &fmt_g23)?;

    // H23:H26 (merged): VLOOKUP(15)
    let fmt_h23 = Format::new()
        .set_font_name("Arial")
        .set_font_size(10.0)
        .set_align(FormatAlign::Center)
        .set_text_wrap();

    let formula_15 = make_vlookup_formula(15, language);
    ws.merge_range(22, 7, 25, 7, "", &fmt_h23)?;
    ws.write_formula_with_format(22, 7, formula_15, &fmt_h23)?;

    // --- ROW 23 (Excel 24): B24:C24 merged ---

    // B24:C24 (merged): VLOOKUP(24) - BOLD
    let fmt_b24 = Format::new()
        .set_font_name("Arial")
        .set_font_size(10.0)
        .set_bold() // ← BOLD
        .set_align(FormatAlign::Center);

    let formula_24 = make_vlookup_formula(24, language);
    ws.merge_range(23, 1, 23, 2, "", &fmt_b24)?;
    ws.write_formula_with_format(23, 1, formula_24, &fmt_b24)?;

    // --- ROW 24 (Excel 25): B25:C25 merged ---

    // B25:C25 (merged): VLOOKUP(10) "Währung"
    let fmt_b25 = Format::new()
        .set_font_name("Arial")
        .set_font_size(10.0)
        .set_align(FormatAlign::Center);

    let formula_10 = make_vlookup_formula(10, language);
    ws.merge_range(24, 1, 24, 2, "", &fmt_b25)?;
    ws.write_formula_with_format(24, 1, formula_10, &fmt_b25)?;

    // --- ROW 25 (Excel 26): B26:C26 merged (leer) ---

    // B26:C26 (merged): Leer
    let fmt_b26 = Format::new()
        .set_font_name("Arial")
        .set_font_size(10.0)
        .set_align(FormatAlign::Center);

    ws.merge_range(25, 1, 25, 2, "", &fmt_b26)?;

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

/// Erstellt eine VLOOKUP-Formel mit gecachtem Text-Ergebnis
fn make_vlookup_formula(index: usize, language: Option<&str>) -> Formula {
    let formula_str = format!(
        r#"=IF($E$2="","",VLOOKUP($E$2,Sprachversionen!$B:$BN,{},FALSE))"#,
        index
    );

    if let Some(text) = lookup_text_string(language, index) {
        Formula::new(&formula_str).set_result(text)
    } else {
        Formula::new(&formula_str)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use rust_xlsxwriter::Workbook;

    #[test]
    fn test_prebody_new_compiles() {
        let mut wb = Workbook::new();
        let ws = wb.add_worksheet();
        let fills = FillColors::new();

        let result = write_prebody_new(ws, None, &fills);
        assert!(result.is_ok());
    }

    #[test]
    fn test_prebody_new_with_language() {
        let mut wb = Workbook::new();
        let ws = wb.add_worksheet();
        let fills = FillColors::new();

        let result = write_prebody_new(ws, Some("deutsch"), &fills);
        assert!(result.is_ok());
    }

    #[test]
    fn test_merged_cells_count() {
        let mut merged_cells = MergedCellRegistry::new();

        // Vertikale Merges (5)
        merged_cells.register_merge(22, 3, 25, 3);
        merged_cells.register_merge(22, 4, 25, 4);
        merged_cells.register_merge(22, 5, 25, 5);
        merged_cells.register_merge(22, 6, 25, 6);
        merged_cells.register_merge(22, 7, 25, 7);

        // Horizontale Merges (3)
        merged_cells.register_merge(23, 1, 23, 2);
        merged_cells.register_merge(24, 1, 24, 2);
        merged_cells.register_merge(25, 1, 25, 2);

        assert_eq!(merged_cells.count(), 8);
    }
}
