//! Right Panel Section - Neues zellbasiertes System
//!
//! Enthält zwei Panels (Zeilen 10-30, Spalten J-V):
//! - Linke Seite: J-O (Cols 9-14)  - Index 1-18
//! - Rechte Seite: Q-V (Cols 16-21) - Index 19-36
//!
//! Struktur pro Panel:
//! - Header (Row 10-12): Merged + Blanks + Formeln (via Registry)
//! - Body (Row 13-30): 18 Zeilen mit Index, Text, Date, Numbers, Calc

use crate::lang::data::TEXT_MATRIX;
use crate::report::api::ReportValues;
use crate::report::format::{BorderManager, BorderSpec, FillColors, MergedCellRegistry};
use rust_xlsxwriter::{Format, FormatAlign, Formula, Worksheet, XlsxError};

/// Schreibt Panel Section mit neuem zellbasiertem System
///
/// # Workflow
///
/// 1. Merged Cells registrieren
/// 2. Borders konfigurieren
/// 3. Header schreiben
/// 4. Body-Rows schreiben (18 Zeilen)
///
/// # Arguments
///
/// * `ws` - Worksheet
/// * `values` - Report Values (für Text-Lookup)
/// * `fills` - Fill Colors
pub fn write_panel_new(
    ws: &mut Worksheet,
    values: &ReportValues,
    fills: &FillColors,
) -> Result<(), XlsxError> {
    // =========================================================================
    // PHASE 1: MERGED CELLS REGISTRIEREN
    // =========================================================================

    let mut merged_cells = MergedCellRegistry::new();
    merged_cells.register_merge(10, 9, 10, 10); // J11:K11 (linker Header)
    merged_cells.register_merge(10, 16, 10, 17); // Q11:R11 (rechter Header)

    // =========================================================================
    // PHASE 2: BORDERS KONFIGURIEREN
    // =========================================================================

    let mut borders = BorderManager::new();
    borders.set_merged_cells(merged_cells.clone());

    // Linkes Panel Box (J11:O30)
    borders.add_range(10, 9, 29, 14, BorderSpec::all_thin());

    // Rechtes Panel Box (Q11:V30)
    borders.add_range(10, 16, 29, 21, BorderSpec::all_thin());

    // =========================================================================
    // PHASE 3: HEADER SCHREIBEN (Row 10-12)
    // =========================================================================

    write_panel_header(ws, &borders, fills, 10, 9)?; // Linker Header
    write_panel_header(ws, &borders, fills, 10, 16)?; // Rechter Header

    // =========================================================================
    // PHASE 4: BODY-ROWS SCHREIBEN (Row 13-30, 18 Zeilen)
    // =========================================================================

    let text_result = evaluate_text_lookup_23(values);
    let f_k = r#"=IF($E$2="","",VLOOKUP($E$2,Sprachversionen!$B:$BN,23,FALSE))"#;

    for i in 0..18u32 {
        let row = 13 + i;
        let left_num = 1 + i;
        let right_num = 19 + i;
        write_panel_body_row(
            ws,
            &borders,
            fills,
            row,
            left_num,
            right_num,
            f_k,
            &text_result,
        )?;
    }

    Ok(())
}

/// Schreibt Panel Header (Row 10-12)
fn write_panel_header(
    ws: &mut Worksheet,
    borders: &BorderManager,
    fills: &FillColors,
    start_row: u32,
    start_col: u16,
) -> Result<(), XlsxError> {
    // Row 10: Merged (J11:K11 oder Q11:R11) + 4 Blanks (L-O oder S-V)
    let fmt_merged = with_borders(
        Format::new()
            .set_font_name("Arial")
            .set_font_size(10.0)
            .set_bold()
            .set_align(FormatAlign::Center),
        borders,
        start_row,
        start_col,
    );

    ws.merge_range(
        start_row,
        start_col,
        start_row,
        start_col + 1,
        "",
        &fmt_merged,
    )?;

    // 4 Blanks rechts (Cols +2 bis +5)
    for offset in 2..=5 {
        let col = start_col + offset;
        let fmt = with_borders(
            Format::new()
                .set_font_name("Arial")
                .set_font_size(10.0)
                .set_align(FormatAlign::Center),
            borders,
            start_row,
            col,
        );
        ws.write_blank(start_row, col, &fmt)?;
    }

    // Row 11: 6 Blanks (alle Spalten)
    for offset in 0..=5 {
        let col = start_col + offset;
        let fmt = with_borders(
            Format::new()
                .set_font_name("Arial")
                .set_font_size(10.0)
                .set_align(FormatAlign::Center),
            borders,
            start_row + 1,
            col,
        );
        ws.write_blank(start_row + 1, col, &fmt)?;
    }

    // Row 12: 2 Blanks (erste 2 Spalten), Rest sind Formeln (via Registry)
    for offset in 0..=1 {
        let col = start_col + offset;
        let fmt = with_borders(
            Format::new()
                .set_font_name("Arial")
                .set_font_size(10.0)
                .set_align(FormatAlign::Center),
            borders,
            start_row + 2,
            col,
        );
        ws.write_blank(start_row + 2, col, &fmt)?;
    }

    Ok(())
}

/// Schreibt eine Panel Body-Zeile (beide Seiten)
fn write_panel_body_row(
    ws: &mut Worksheet,
    borders: &BorderManager,
    fills: &FillColors,
    row: u32,
    left_num: u32,
    right_num: u32,
    f_k: &str,
    text_result: &str,
) -> Result<(), XlsxError> {
    // === LINKE SEITE (J-O, Cols 9-14) ===
    write_panel_side(ws, borders, fills, row, 9, left_num, f_k, text_result)?;

    // === RECHTE SEITE (Q-V, Cols 16-21) ===
    write_panel_side(ws, borders, fills, row, 16, right_num, f_k, text_result)?;

    Ok(())
}

/// Schreibt eine Seite einer Panel-Zeile (6 Spalten)
fn write_panel_side(
    ws: &mut Worksheet,
    borders: &BorderManager,
    fills: &FillColors,
    row: u32,
    start_col: u16,
    num: u32,
    f_k: &str,
    text_result: &str,
) -> Result<(), XlsxError> {
    // Spalte 0: Index (z.B. "1. ")
    let fmt_index = with_borders(
        Format::new()
            .set_font_name("Arial")
            .set_font_size(10.0)
            .set_align(FormatAlign::Right)
            .set_num_format("@"),
        borders,
        row,
        start_col,
    );
    ws.write_string_with_format(row, start_col, &format!("{}. ", num), &fmt_index)?;

    // Spalte 1: Text (VLOOKUP-Formel, locked)
    let fmt_text = with_borders(
        Format::new()
            .set_font_name("Arial")
            .set_font_size(10.0)
            .set_align(FormatAlign::Left)
            .set_num_format("@"),
        borders,
        row,
        start_col + 1,
    );
    let formula_text = Formula::new(f_k).set_result(text_result);
    ws.write_formula_with_format(row, start_col + 1, formula_text, &fmt_text)?;

    // Spalte 2: Date (blank, gelb, unlocked)
    let fmt_date = with_borders(
        Format::new()
            .set_font_name("Arial")
            .set_font_size(10.0)
            .set_background_color(fills.input)
            .set_align(FormatAlign::Center)
            .set_unlocked()
            .set_num_format("dd.mm.yyyy"),
        borders,
        row,
        start_col + 2,
    );
    ws.write_blank(row, start_col + 2, &fmt_date)?;

    // Spalte 3: Number (blank, gelb, unlocked)
    let fmt_num1 = with_borders(
        Format::new()
            .set_font_name("Arial")
            .set_font_size(10.0)
            .set_background_color(fills.input)
            .set_align(FormatAlign::Right)
            .set_unlocked()
            .set_num_format("#,##0.00"),
        borders,
        row,
        start_col + 3,
    );
    ws.write_blank(row, start_col + 3, &fmt_num1)?;

    // Spalte 4: Number (blank, gelb, unlocked)
    let fmt_num2 = with_borders(
        Format::new()
            .set_font_name("Arial")
            .set_font_size(10.0)
            .set_background_color(fills.input)
            .set_align(FormatAlign::Right)
            .set_unlocked()
            .set_num_format("#,##0.00"),
        borders,
        row,
        start_col + 4,
    );
    ws.write_blank(row, start_col + 4, &fmt_num2)?;

    // Spalte 5: Calc (Formel, locked, percentage)
    let col3_letter = col_to_letter(start_col + 3);
    let col4_letter = col_to_letter(start_col + 4);
    let f_calc = format!(
        "=IF({}{}=\"\",\"\",{}{}/{}{})",
        col3_letter,
        row + 1,
        col4_letter,
        row + 1,
        col3_letter,
        row + 1
    );
    let fmt_calc = with_borders(
        Format::new()
            .set_font_name("Arial")
            .set_font_size(10.0)
            .set_align(FormatAlign::Center)
            .set_num_format("0.00%"),
        borders,
        row,
        start_col + 5,
    );
    let formula_calc = Formula::new(&f_calc).set_result("");
    ws.write_formula_with_format(row, start_col + 5, formula_calc, &fmt_calc)?;

    Ok(())
}

/// Konvertiert Spalten-Index zu Buchstabe (z.B. 0 -> A, 9 -> J)
fn col_to_letter(col: u16) -> String {
    let mut result = String::new();
    let mut n = col + 1; // 0-indexed → 1-indexed

    while n > 0 {
        n -= 1;
        result.insert(0, (b'A' + (n % 26) as u8) as char);
        n /= 26;
    }

    result
}

/// Evaluiert Text-Lookup für Index 23 (Panel K/R Spalten)
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
            .and_then(|row| row.get(22)) // Index 23 in Excel = Index 22 in 0-basiert
            .map(|s| s.to_string())
            .unwrap_or_default(),
        None => String::new(),
    }
}

/// Helper-Funktion für Border-Merge
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
    fn test_panel_new_compiles() {
        let mut wb = Workbook::new();
        let ws = wb.add_worksheet();
        let fills = FillColors::new();
        let values = ReportValues::new().with_language("deutsch");

        let result = write_panel_new(ws, &values, &fills);
        assert!(result.is_ok());
    }

    #[test]
    fn test_col_to_letter() {
        assert_eq!(col_to_letter(0), "A");
        assert_eq!(col_to_letter(9), "J");
        assert_eq!(col_to_letter(12), "M");
        assert_eq!(col_to_letter(14), "O");
        assert_eq!(col_to_letter(16), "Q");
        assert_eq!(col_to_letter(25), "Z");
        assert_eq!(col_to_letter(26), "AA");
    }
}
