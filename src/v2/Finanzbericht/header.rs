use super::styles::ReportStyles;
use rust_xlsxwriter::{DataValidation, Format, FormatAlign, Worksheet, XlsxError};

struct LocalStyles {
    fmt_top_med: Format,
    fmt_top_right_med: Format,
    fmt_d5: Format,
    fmt_d6: Format,
    fmt_row7_base: Format,
    fmt_row7_date: Format,
    fmt_th_b: Format,
    fmt_th_d: Format,
    fmt_th_d_bold: Format,
    fmt_th_h: Format,
    fmt_th_side: Format,
    fmt_th_side_bold: Format,
    fmt_th_c_top: Format,
    fmt_th_bot_side: Format,
    fmt_th_bot_right: Format,
    orange_dotted: Format,
    value_dotted: Format,
    input_dotted: Format,
}

impl LocalStyles {
    fn new(s: &ReportStyles) -> Self {
        let fmt_top_med = s.base.clone().set_border_top(s.border_medium);
        let fmt_top_right_med = s
            .base
            .clone()
            .set_border_top(s.border_medium)
            .set_border_right(s.border_medium);

        let fmt_d5 = s
            .base
            .clone()
            .set_background_color(s.fill_input)
            .set_border_bottom(s.border_thin);

        let fmt_d6 = s
            .base
            .clone()
            .set_background_color(s.fill_input)
            .set_align(FormatAlign::Left)
            .set_align(FormatAlign::Top)
            .set_text_wrap()
            .set_border_bottom(s.border_thin);

        let fmt_row7_base = s.base.clone().set_border_bottom(s.border_thin);
        let fmt_row7_date = fmt_row7_base
            .clone()
            .set_background_color(s.fill_input)
            .set_num_format("mm-dd-yy")
            .set_align(FormatAlign::Left);

        let fmt_th_b = s.table_header_base.clone().set_border_left(s.border_medium);

        let fmt_th_c_top = s
            .table_header_base
            .clone()
            .set_border_right(s.border_medium);

        let fmt_th_d = s
            .table_header_base
            .clone()
            .set_border_left(s.border_thin)
            .set_border_right(s.border_thin);

        let fmt_th_d_bold = fmt_th_d.clone().set_bold();

        let fmt_th_h = s
            .table_header_base
            .clone()
            .set_border_left(s.border_thin)
            .set_border_right(s.border_medium);

        let fmt_th_side = s
            .base
            .clone()
            .set_align(FormatAlign::Center)
            .set_align(FormatAlign::VerticalCenter)
            .set_border_left(s.border_medium)
            .set_border_right(s.border_medium);

        let fmt_th_side_bold = fmt_th_side.clone().set_bold();

        let fmt_th_bot_side = s
            .base
            .clone()
            .set_align(FormatAlign::Center)
            .set_align(FormatAlign::VerticalCenter)
            .set_border_left(s.border_medium)
            .set_border_bottom(s.border_thin);

        let fmt_th_bot_right = s
            .base
            .clone()
            .set_border_bottom(s.border_thin)
            .set_border_right(s.border_medium);

        let orange_dotted = s
            .base
            .clone()
            .set_background_color(s.fl_orange)
            .set_border(s.border_dotted);
        let value_dotted = s
            .base
            .clone()
            .set_background_color(s.fill_value)
            .set_border(s.border_dotted);
        let input_dotted = s
            .base
            .clone()
            .set_background_color(s.fill_input)
            .set_border(s.border_dotted);

        Self {
            fmt_top_med,
            fmt_top_right_med,
            fmt_d5,
            fmt_d6,
            fmt_row7_base,
            fmt_row7_date,
            fmt_th_b,
            fmt_th_d,
            fmt_th_d_bold,
            fmt_th_h,
            fmt_th_side,
            fmt_th_side_bold,
            fmt_th_c_top,
            fmt_th_bot_side,
            fmt_th_bot_right,
            orange_dotted,
            value_dotted,
            input_dotted,
        }
    }
}

pub fn write_header(
    ws: &mut Worksheet,
    styles: &ReportStyles,
    suffix: &str,
    lang_val: &str,
) -> Result<(), XlsxError> {
    let local_styles = LocalStyles::new(styles);

    set_formatting(ws, styles, &local_styles)?;
    set_values(ws, styles, suffix, lang_val)?;
    set_formulas(ws, styles, &local_styles)?;

    Ok(())
}

fn set_formatting(
    ws: &mut Worksheet,
    styles: &ReportStyles,
    ls: &LocalStyles,
) -> Result<(), XlsxError> {
    // --- Row 0 ---
    // B1:C1 Merged
    ws.merge_range(0, 1, 0, 2, "", &styles.header_label)?;
    // K1:O1 (Top Border Medium)
    for col in 10..=13 {
        ws.write_blank(0, col, &ls.fmt_top_med)?;
    }
    // O1 (Top Right Corner)
    ws.write_blank(0, 14, &ls.fmt_top_right_med)?;

    // --- Row 1 ---
    // B2:C2 Merged
    ws.merge_range(1, 1, 1, 2, "", &styles.header_suffix)?;
    // J2:O3 Merged
    ws.merge_range(1, 9, 2, 14, "", &styles.header_right_box_body)?;

    // --- Row 2 ---
    // B3:C3 Merged
    ws.merge_range(2, 1, 2, 2, "", &styles.header_label)?;

    // --- Row 3 ---
    // J4:O4 Merged
    ws.merge_range(3, 9, 3, 14, "", &styles.link_style)?;

    // --- Row 4 ---
    // B5:C5 Merged
    ws.merge_range(4, 1, 4, 2, "", &styles.left_center)?;
    // D5
    ws.write_blank(4, 3, &ls.fmt_d5)?;

    // --- Row 5 ---
    // B6:C7 Merged
    ws.merge_range(5, 1, 6, 2, "", &styles.left_center)?;
    // D6:H7 Merged
    ws.merge_range(5, 3, 6, 7, "", &ls.fmt_d6)?;

    // --- Row 7 (Excel 8) ---
    // B8:C8 Merged
    ws.merge_range(7, 1, 7, 2, "", &styles.left_center)?;
    // G8:H8 merged
    ws.merge_range(7, 6, 7, 7, "", &ls.fmt_row7_date)?;
    // E8 (Blank with style)
    ws.write_blank(7, 4, &ls.fmt_row7_date)?;
    // J7 (Excel) -> Row 6 in 0-index.
    ws.write_blank(6, 9, &ls.orange_dotted)?;
    // J8 (Excel) -> Row 7 in 0-index.
    ws.write_blank(7, 9, &ls.value_dotted)?;

    // --- Row 8 (Excel 9) ---
    // B9:C9 Merged
    ws.merge_range(8, 1, 8, 2, "", &styles.left_center)?;
    // G9:H9 Merged
    ws.merge_range(8, 6, 8, 7, "", &ls.fmt_row7_date)?;
    // J9 (Excel) -> Row 8 in 0-index.
    ws.write_blank(8, 9, &ls.input_dotted)?;

    // --- Row 10 (Table Header) ---
    // B11, C11 borders
    ws.write_blank(10, 1, &ls.fmt_th_b)?;
    ws.write_blank(10, 2, &ls.fmt_th_c_top)?;

    // J11:K11 Merged
    ws.merge_range(10, 9, 10, 10, "", &styles.left_center_bold)?;
    // Q11:R11 Merged
    ws.merge_range(10, 16, 10, 17, "", &styles.left_center_bold)?;

    // Vertical Merges D11-H13
    ws.merge_range(10, 3, 13, 3, "", &ls.fmt_th_d)?;
    ws.merge_range(10, 4, 13, 4, "", &ls.fmt_th_d_bold)?;
    ws.merge_range(10, 5, 13, 5, "", &ls.fmt_th_d)?;
    ws.merge_range(10, 6, 13, 6, "", &ls.fmt_th_d)?;
    ws.merge_range(10, 7, 13, 7, "", &ls.fmt_th_h)?;

    // --- Row 11 ---
    // B12:C12
    ws.merge_range(11, 1, 11, 2, "", &ls.fmt_th_side_bold)?;

    // --- Row 12 ---
    // B13:C13
    ws.merge_range(12, 1, 12, 2, "", &ls.fmt_th_side)?;

    // --- Row 13 ---
    // B14
    ws.write_blank(13, 1, &ls.fmt_th_bot_side)?;
    // C14
    ws.write_blank(13, 2, &ls.fmt_th_bot_right)?;

    Ok(())
}

fn set_values(
    ws: &mut Worksheet,
    styles: &ReportStyles,
    suffix: &str,
    lang_val: &str,
) -> Result<(), XlsxError> {
    // --- Row 1 ---
    // B2: Suffix (Merged in fmt pass, writing to B2)
    // Note: Writing string with format to preserve the merge format?
    // Actually, for merged ranges, if we write to top-left, we must supply format if we want it to look right.
    ws.write_string_with_format(1, 1, suffix, &styles.header_suffix)?;

    // E2: Language Input
    ws.write_string_with_format(1, 4, lang_val, &styles.input_orange_dashed)?;
    let validation = DataValidation::new().allow_list_formula("=Sprachversionen!$B$1:$B$5".into());
    ws.add_data_validation(1, 4, 1, 4, &validation)?;

    // --- Row 12 ---
    // M13, T13: Euro
    ws.write_string_with_format(12, 12, "Euro", &styles.center_center_bold)?;
    ws.write_string_with_format(12, 19, "Euro", &styles.center_center_bold)?;

    Ok(())
}

fn set_formulas(
    ws: &mut Worksheet,
    styles: &ReportStyles,
    ls: &LocalStyles,
) -> Result<(), XlsxError> {
    // --- Row 0 ---
    ws.write_formula_with_format(
        0,
        1,
        "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,2,FALSE))",
        &styles.header_label,
    )?;
    ws.write_formula_with_format(
        0,
        9,
        "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,60,FALSE))",
        &styles.header_right_box_top,
    )?;

    // --- Row 1 ---
    ws.write_formula_with_format(
        1,
        3,
        "=IF($E$2=\"\",\"Chose your language\",VLOOKUP($E$2,Sprachversionen!$B:$BN,27,FALSE))",
        &styles.left_center,
    )?;
    ws.write_formula_with_format(
        1,
        9,
        "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,61,FALSE))",
        &styles.header_right_box_body,
    )?;

    // --- Row 2 ---
    ws.write_formula_with_format(
        2,
        1,
        "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,3,FALSE))",
        &styles.header_label,
    )?;
    ws.write_formula_with_format(
        2,
        3,
        "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,28,FALSE))",
        &styles.left_center,
    )?;
    ws.write_formula_with_format(
        2,
        4,
        "=IF($E$2=\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,28,FALSE),E3)",
        &styles.input_orange_dashed,
    )?;

    // --- Row 3 ---
    ws.write_formula_with_format(
        3,
        9,
        "=HYPERLINK(VLOOKUP($E$2,Sprachversionen!$B:$BN,62,FALSE))",
        &styles.link_style,
    )?;

    // --- Row 4 ---
    ws.write_formula_with_format(
        4,
        1,
        "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,4,FALSE))",
        &styles.left_center,
    )?;

    // --- Row 5 ---
    ws.write_formula_with_format(
        5,
        1,
        "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,26,FALSE))",
        &styles.left_center,
    )?;

    // --- Row 7 ---
    ws.write_formula_with_format(
        7,
        1,
        "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,5,FALSE))",
        &styles.left_center,
    )?;
    ws.write_formula_with_format(
        7,
        3,
        "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,7,FALSE))",
        &ls.fmt_row7_base,
    )?;
    ws.write_formula_with_format(
        7,
        5,
        "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,8,FALSE))",
        &ls.fmt_row7_base,
    )?;
    // K8
    ws.write_formula(
        7,
        10,
        "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,59,FALSE))",
    )?;

    // --- Row 8 ---
    ws.write_formula_with_format(
        8,
        1,
        "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,6,FALSE))",
        &styles.left_center,
    )?;
    ws.write_formula_with_format(8, 3, "=D8", &ls.fmt_row7_base)?;
    ws.write_blank(8, 4, &ls.fmt_row7_date)?;
    ws.write_formula_with_format(8, 5, "=F8", &ls.fmt_row7_base)?;
    ws.write_blank(8, 6, &ls.fmt_row7_date)?;

    // --- Row 10 (Header Table) ---
    // J11, Q11
    ws.write_formula_with_format(10, 9, "=B18", &styles.left_center_bold)?;
    ws.write_formula_with_format(10, 16, "=B18", &styles.left_center_bold)?;

    // Vertical columns
    ws.write_formula_with_format(
        10,
        3,
        "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,11,FALSE))",
        &ls.fmt_th_d,
    )?;
    ws.write_formula_with_format(
        10,
        4,
        "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,12,FALSE))",
        &ls.fmt_th_d_bold,
    )?;
    ws.write_formula_with_format(
        10,
        5,
        "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,13,FALSE))",
        &ls.fmt_th_d,
    )?;
    ws.write_formula_with_format(
        10,
        6,
        "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,14,FALSE))",
        &ls.fmt_th_d,
    )?;
    ws.write_formula_with_format(
        10,
        7,
        "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,15,FALSE))",
        &ls.fmt_th_h,
    )?;

    // --- Row 11 ---
    ws.write_formula_with_format(
        11,
        1,
        "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,9,FALSE))",
        &ls.fmt_th_side_bold,
    )?;

    // --- Row 12 ---
    ws.write_formula_with_format(
        12,
        1,
        "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,10,FALSE))",
        &ls.fmt_th_side,
    )?;

    // L13, N13, O13, S13, U13, V13
    ws.write_formula_with_format(
        12,
        11,
        "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,22,FALSE))",
        &styles.center_center_bold,
    )?;
    ws.write_formula_with_format(
        12,
        13,
        "=IF(E3=\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,28,FALSE),E3)",
        &styles.center_center_bold,
    )?;
    ws.write_formula_with_format(
        12,
        14,
        "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,58,FALSE))",
        &styles.center_center_bold,
    )?;
    ws.write_formula_with_format(
        12,
        18,
        "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,22,FALSE))",
        &styles.center_center_bold,
    )?;
    ws.write_formula_with_format(
        12,
        20,
        "=IF(E3=\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,28,FALSE),E3)",
        &styles.center_center_bold,
    )?;
    ws.write_formula_with_format(
        12,
        21,
        "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,58,FALSE))",
        &styles.center_center_bold,
    )?;

    Ok(())
}
