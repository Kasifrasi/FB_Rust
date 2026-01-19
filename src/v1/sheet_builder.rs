use rust_xlsxwriter::{
    Color, DataValidation, Format, FormatAlign, FormatBorder, Workbook, Worksheet, XlsxError,
};

/// Clone a base format and override individual borders.
pub fn with_border(
    base: &Format,
    left: Option<FormatBorder>,
    right: Option<FormatBorder>,
    top: Option<FormatBorder>,
    bottom: Option<FormatBorder>,
) -> Format {
    let mut fmt = base.clone();
    if let Some(b) = left {
        fmt = fmt.set_border_left(b);
    }
    if let Some(b) = right {
        fmt = fmt.set_border_right(b);
    }
    if let Some(b) = top {
        fmt = fmt.set_border_top(b);
    }
    if let Some(b) = bottom {
        fmt = fmt.set_border_bottom(b);
    }
    fmt
}

/// Convert zero-based row/col to an Excel address like "B27".
pub fn addr(col: char, row: u32) -> String {
    format!("{}{}", col, row + 1)
}

/// Helper to hide a column by its Excel letter.
pub fn hide_column(ws: &mut Worksheet, col: char) {
    let col_idx = match col {
        'A'..='Z' => col as u16 - 'A' as u16,
        _ => return,
    };
    let _ = ws.set_column_hidden(col_idx);
}

pub fn recreate_sheet(
    workbook: &mut Workbook,
    sheet_name: &str,
    suffix: &str,
    lang_val: &str,
) -> Result<(), XlsxError> {
    let ws = workbook.add_worksheet().set_name(sheet_name)?;

    // --- Page Setup ---
    ws.set_zoom(85);
    ws.set_screen_gridlines(false);

    // --- Columns ---
    // A=0
    ws.set_column_width(0, 3.285)?;
    ws.set_column_width(1, 4.57)?;
    ws.set_column_width(2, 55.85)?;
    ws.set_column_width(3, 20.0)?;
    ws.set_column_width(4, 18.57)?;
    ws.set_column_width(5, 18.71)?;
    ws.set_column_width(6, 11.71)?;
    ws.set_column_width(7, 18.71)?;
    ws.set_column_width(8, 4.0)?; // I
    ws.set_column_width(9, 4.0)?; // J
    ws.set_column_width(10, 19.14)?; // K
    ws.set_column_width(11, 10.85)?; // L
    ws.set_column_width(12, 15.0)?; // M
    ws.set_column_width(13, 15.0)?; // N
    ws.set_column_width(14, 15.0)?; // O
    ws.set_column_width(15, 5.0)?; // P
    ws.set_column_width(16, 4.0)?; // Q
    ws.set_column_width(17, 19.14)?; // R
    ws.set_column_width(18, 10.85)?; // S
    ws.set_column_width(19, 15.0)?; // T
    ws.set_column_width(20, 15.0)?; // U
    ws.set_column_width(21, 15.0)?; // V
    ws.set_column_width(22, 4.0)?; // W
    ws.set_column_width(23, 36.71)?; // X
    ws.set_column_width(24, 10.85)?; // Y

    // Grouping Q-V (16-21)
    // Note: rust_xlsxwriter sets level via set_column_options, but simpler just to set widths.
    // If we wanted to group, we'd need ColOptions. For now, assume simple widths are enough or add grouping later if needed.

    // --- Colors & Common Formats ---
    let fill_input = Color::RGB(0xFFFAE5);
    let fill_value = Color::RGB(0xF2F2F2);
    let fill_summary = Color::RGB(0xD9D9D9);
    let fl_orange = Color::RGB(0xFFCC00);
    let blue_link = Color::RGB(0x0563C1);

    // Base Alignments
    let left_center = Format::new()
        .set_align(FormatAlign::Left)
        .set_align(FormatAlign::VerticalCenter);
    let center_center = Format::new()
        .set_align(FormatAlign::Center)
        .set_align(FormatAlign::VerticalCenter);
    // let right_center = Format::new()
    //     .set_align(FormatAlign::Right)
    //     .set_align(FormatAlign::VerticalCenter);

    // Base Borders
    let border_thin = FormatBorder::Thin;
    let border_medium = FormatBorder::Medium;
    let border_dotted = FormatBorder::Dotted;
    let border_dashed = FormatBorder::Dashed;

    // --- Row 0 (Header Top) ---
    // B1:C1 Merged: VLOOKUP(..., 2)
    let fmt_header_label = left_center.clone().set_bold();
    ws.merge_range(0, 1, 0, 2, "", &fmt_header_label)?;
    ws.write_formula_with_format(
        0,
        1,
        "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,2,FALSE))",
        &fmt_header_label,
    )?;

    // J1: Part of Header Right Box (Top Left Corner)
    // Box J1:O4 has Medium Outline. J1 is Top/Left Medium.
    let fmt_j1 = Format::new()
        .set_bold()
        .set_border_top(border_medium)
        .set_border_left(border_medium)
        .set_align(FormatAlign::Left);
    ws.write_formula_with_format(
        0,
        9,
        "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,60,FALSE))",
        &fmt_j1,
    )?;

    // K1:O1 (Top Border Medium)
    let fmt_top_med = Format::new().set_border_top(border_medium);
    for col in 10..=13 {
        ws.write_blank(0, col, &fmt_top_med)?;
    }
    // O1 (Top Right Corner)
    let fmt_top_right_med = Format::new()
        .set_border_top(border_medium)
        .set_border_right(border_medium);
    ws.write_blank(0, 14, &fmt_top_right_med)?;

    // --- Row 1 ---
    // B2:C2 Merged: Suffix
    let fmt_suffix = left_center.clone().set_font_color(Color::White);
    ws.merge_range(1, 1, 1, 2, suffix, &fmt_suffix)?;

    // D2: Language Label Formula
    ws.write_formula_with_format(
        1,
        3,
        "=IF($E$2=\"\",\"Chose your language\",VLOOKUP($E$2,Sprachversionen!$B:$BN,27,FALSE))",
        &left_center,
    )?;

    // E2: Language Input (Validation)
    let fmt_lang_input = center_center
        .clone()
        .set_background_color(fl_orange)
        .set_border_bottom(border_dashed);
    ws.write_string_with_format(1, 4, lang_val, &fmt_lang_input)?;
    let validation = DataValidation::new().allow_list_formula("=Sprachversionen!$B$1:$B$5".into());
    ws.add_data_validation(1, 4, 1, 4, &validation)?;

    // J2:O3 Merged. Header Right Box Body.
    // J2 has Left Medium. O3 has Right Medium.
    let fmt_header_desc = Format::new()
        .set_align(FormatAlign::Left)
        .set_align(FormatAlign::Top)
        .set_text_wrap()
        .set_border_left(border_medium)
        .set_border_right(border_medium);

    ws.merge_range(1, 9, 2, 14, "", &fmt_header_desc)?;
    ws.write_formula_with_format(
        1,
        9,
        "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,61,FALSE))",
        &fmt_header_desc,
    )?;

    // --- Row 2 ---
    // B3:C3 Merged
    ws.merge_range(2, 1, 2, 2, "", &fmt_header_label)?;
    ws.write_formula_with_format(
        2,
        1,
        "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,3,FALSE))",
        &fmt_header_label,
    )?;

    // D3: Currency Label
    ws.write_formula_with_format(
        2,
        3,
        "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,28,FALSE))",
        &left_center,
    )?;

    // E3: Currency Input
    // Fill Orange, Bottom Dashed (same as E2)
    ws.write_formula_with_format(
        2,
        4,
        "=IF($E$2=\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,28,FALSE),E3)",
        &fmt_lang_input,
    )?;

    // --- Row 3 ---
    ws.set_row_height(3, 12.75)?;

    // J4:O4 Merged. Link.
    let fmt_link = Format::new()
        .set_align(FormatAlign::Left)
        .set_font_color(blue_link)
        .set_border_left(border_medium)
        .set_border_right(border_medium)
        .set_border_bottom(border_medium);

    ws.merge_range(3, 9, 3, 14, "", &fmt_link)?;
    ws.write_formula_with_format(
        3,
        9,
        "=HYPERLINK(VLOOKUP($E$2,Sprachversionen!$B:$BN,62,FALSE))",
        &fmt_link,
    )?;

    // --- Row 4 ---
    ws.set_row_height(4, 12.0)?;
    // B5:C5
    ws.merge_range(4, 1, 4, 2, "", &left_center)?;
    ws.write_formula_with_format(
        4,
        1,
        "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,4,FALSE))",
        &left_center,
    )?;

    // D5:D6 Border Bottom Thin. D5 also fill input?
    let fmt_d5 = Format::new()
        .set_background_color(fill_input)
        .set_border_bottom(border_thin);
    ws.write_blank(4, 3, &fmt_d5)?;

    // --- Row 5 ---
    // B6:C7 Merged
    ws.merge_range(5, 1, 6, 2, "", &left_center)?;
    ws.write_formula_with_format(
        5,
        1,
        "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,26,FALSE))",
        &left_center,
    )?;

    // D6:H7 Merged? D6 Left Wrap.
    let fmt_d6 = Format::new()
        .set_background_color(fill_input)
        .set_align(FormatAlign::Left)
        .set_align(FormatAlign::Top)
        .set_text_wrap()
        .set_border_bottom(border_thin);

    ws.merge_range(5, 3, 6, 7, "", &fmt_d6)?;

    // --- Row 7 ---
    // B8:C8
    ws.merge_range(7, 1, 7, 2, "", &left_center)?;
    ws.write_formula_with_format(
        7,
        1,
        "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,5,FALSE))",
        &left_center,
    )?;

    // D8, F8 Vlookups. E8, G8 Dates.
    let fmt_row7_base = Format::new().set_border_bottom(border_thin);
    let fmt_row7_date = fmt_row7_base
        .clone()
        .set_background_color(fill_input)
        .set_num_format("mm-dd-yy")
        .set_align(FormatAlign::Left);

    ws.write_formula_with_format(
        7,
        3,
        "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,7,FALSE))",
        &fmt_row7_base,
    )?;
    ws.write_blank(7, 4, &fmt_row7_date)?; // E8
    ws.write_formula_with_format(
        7,
        5,
        "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,8,FALSE))",
        &fmt_row7_base,
    )?;
    ws.write_blank(7, 6, &fmt_row7_date)?; // G8
                                           // G8:H8 merged
    ws.merge_range(7, 6, 7, 7, "", &fmt_row7_date)?;

    // K8
    ws.write_formula(
        7,
        10,
        "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,59,FALSE))",
    )?;

    // J8 (J7 in 0-based is row 6)
    // Row 6 (Excel 7): J7. Orange.
    ws.write_blank(
        6,
        9,
        &Format::new()
            .set_background_color(fl_orange)
            .set_border(border_dotted),
    )?;

    // Row 7 (Excel 8): J8. Value fill.
    ws.write_blank(
        7,
        9,
        &Format::new()
            .set_background_color(fill_value)
            .set_border(border_dotted),
    )?;

    // --- Row 8 ---
    // B9:C9
    ws.merge_range(8, 1, 8, 2, "", &left_center)?;
    ws.write_formula_with_format(
        8,
        1,
        "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,6,FALSE))",
        &left_center,
    )?;

    // D9=D8, F9=F8. E9, G9 Dates.
    ws.write_formula_with_format(8, 3, "=D8", &fmt_row7_base)?;
    ws.write_formula_with_format(8, 4, "=E8", &fmt_row7_date)?; // E9
    ws.write_formula_with_format(8, 5, "=F8", &fmt_row7_base)?;
    ws.write_blank(8, 6, &fmt_row7_date)?; // G9
    ws.merge_range(8, 6, 8, 7, "", &fmt_row7_date)?;
    ws.write_formula_with_format(8, 6, "=G8", &fmt_row7_date)?;

    // J9. Input fill.
    ws.write_blank(
        8,
        9,
        &Format::new()
            .set_background_color(fill_input)
            .set_border(border_dotted),
    )?;

    // --- Row 9 ---
    ws.set_row_height(9, 13.5)?;

    // --- Row 10 (Table Header) ---
    ws.set_row_height(10, 12.6)?;

    let fmt_th_b = Format::new()
        .set_bold()
        .set_align(FormatAlign::Center)
        .set_align(FormatAlign::VerticalCenter)
        .set_text_wrap()
        .set_border_top(border_medium)
        .set_border_left(border_medium)
        .set_border_right(border_thin);

    let fmt_th_d = Format::new()
        .set_align(FormatAlign::Center)
        .set_align(FormatAlign::VerticalCenter)
        .set_text_wrap()
        .set_border_top(border_medium)
        .set_border_left(border_thin)
        .set_border_right(border_thin);

    let fmt_th_h = Format::new()
        .set_align(FormatAlign::Center)
        .set_align(FormatAlign::VerticalCenter)
        .set_text_wrap()
        .set_border_top(border_medium)
        .set_border_left(border_thin)
        .set_border_right(border_medium);

    // B11 Vlookup
    ws.write_formula_with_format(
        10,
        1,
        "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,11,FALSE))",
        &fmt_th_b,
    )?;
    // J11:K11 Merged (Right Panel Header)
    ws.merge_range(
        10,
        9,
        10,
        10,
        "",
        &Format::new().set_bold().set_align(FormatAlign::Left),
    )?;
    ws.write_formula_with_format(
        10,
        9,
        "=B18",
        &Format::new().set_bold().set_align(FormatAlign::Left),
    )?;

    // Q11:R11 Merged
    ws.merge_range(
        10,
        16,
        10,
        17,
        "",
        &Format::new().set_bold().set_align(FormatAlign::Left),
    )?;
    ws.write_formula_with_format(
        10,
        16,
        "=B18",
        &Format::new().set_bold().set_align(FormatAlign::Left),
    )?;

    // D11..H11 merged vertically to row 13
    ws.merge_range(10, 3, 13, 3, "", &fmt_th_d)?;
    ws.write_formula_with_format(
        10,
        3,
        "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,11,FALSE))",
        &fmt_th_d,
    )?;

    ws.merge_range(10, 4, 13, 4, "", &fmt_th_d.clone().set_bold())?;
    ws.write_formula_with_format(
        10,
        4,
        "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,12,FALSE))",
        &fmt_th_d.clone().set_bold(),
    )?; // E11 Bold

    ws.merge_range(10, 5, 13, 5, "", &fmt_th_d)?;
    ws.write_formula_with_format(
        10,
        5,
        "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,13,FALSE))",
        &fmt_th_d,
    )?;

    ws.merge_range(10, 6, 13, 6, "", &fmt_th_d)?;
    ws.write_formula_with_format(
        10,
        6,
        "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,14,FALSE))",
        &fmt_th_d,
    )?;

    ws.merge_range(10, 7, 13, 7, "", &fmt_th_h)?;
    ws.write_formula_with_format(
        10,
        7,
        "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,15,FALSE))",
        &fmt_th_h,
    )?;

    // --- Rows 11, 12, 13 (Rest of Header) ---
    // B12, B13
    let fmt_th_side = Format::new()
        .set_align(FormatAlign::Center)
        .set_align(FormatAlign::VerticalCenter)
        .set_border_left(border_medium)
        .set_border_right(border_thin);

    // Row 11 (Excel 12)
    ws.merge_range(11, 1, 11, 2, "", &fmt_th_side.clone().set_bold())?;
    ws.write_formula_with_format(
        11,
        1,
        "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,9,FALSE))",
        &fmt_th_side.clone().set_bold(),
    )?;

    // Row 12 (Excel 13)
    ws.merge_range(12, 1, 12, 2, "", &fmt_th_side)?;
    ws.write_formula_with_format(
        12,
        1,
        "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,10,FALSE))",
        &fmt_th_side,
    )?;

    // Right Panel Headers in Row 12 (Excel 13)
    let fmt_rp_header = Format::new().set_bold().set_align(FormatAlign::Center);
    ws.write_string_with_format(12, 12, "Euro", &fmt_rp_header)?; // M13
    ws.write_string_with_format(12, 19, "Euro", &fmt_rp_header)?; // T13

    ws.write_formula_with_format(
        12,
        11,
        "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,22,FALSE))",
        &fmt_rp_header,
    )?; // L13
    ws.write_formula_with_format(
        12,
        13,
        "=IF(E3=\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,28,FALSE),E3)",
        &fmt_rp_header,
    )?; // N13
    ws.write_formula_with_format(
        12,
        14,
        "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,58,FALSE))",
        &fmt_rp_header,
    )?; // O13

    ws.write_formula_with_format(
        12,
        18,
        "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,22,FALSE))",
        &fmt_rp_header,
    )?; // S13
    ws.write_formula_with_format(
        12,
        20,
        "=IF(E3=\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,28,FALSE),E3)",
        &fmt_rp_header,
    )?; // U13
    ws.write_formula_with_format(
        12,
        21,
        "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,58,FALSE))",
        &fmt_rp_header,
    )?; // V13

    // --- Row 13 (Excel 14) Bottom of Header ---
    let fmt_th_bot = Format::new().set_border_bottom(border_thin);
    ws.write_blank(13, 1, &fmt_th_side.clone().set_border_bottom(border_thin))?; // B14
    ws.write_blank(13, 2, &fmt_th_bot.clone().set_border_right(border_thin))?; // C14

    // --- Body (Rows 14-19) ---
    // Indexes 14 to 19 (Excel 15-20)

    let fmt_body_num = Format::new()
        .set_num_format("#,##0.00")
        .set_border(border_thin);
    let fmt_body_pct = Format::new().set_num_format("0%").set_border(border_thin);

    let fmt_val_bg = fmt_body_num.clone().set_background_color(fill_value);
    let fmt_inp_bg = fmt_body_num.clone().set_background_color(fill_input);
    let fmt_inp_txt = Format::new()
        .set_background_color(fill_input)
        .set_border(border_thin);

    let fmt_lbl = Format::new()
        .set_align(FormatAlign::Left)
        .set_bold()
        .set_border_left(border_medium) // Outer Left
        .set_border_right(border_thin) // Inner Right
        .set_border_top(border_thin)
        .set_border_bottom(border_thin);

    let vlookup_indices = [16, 17, 18, 19, 20]; // For B15..B19

    for i in 0..5 {
        let row = 14 + i as u32;
        let v_idx = vlookup_indices[i];

        // B Col
        let f_b = format!(
            "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,{},FALSE))",
            v_idx
        );
        ws.merge_range(row, 1, row, 2, "", &fmt_lbl)?;
        ws.write_formula_with_format(row, 1, f_b.as_str(), &fmt_lbl)?;

        // D: Value
        ws.write_number_with_format(row, 3, 0.0, &fmt_val_bg)?;
        // E: Input
        ws.write_blank(row, 4, &fmt_inp_bg)?;
        // F: Calc
        ws.write_blank(row, 5, &fmt_body_num)?;
        // G: %
        if row >= 15 {
            // G16:G20
            let f = format!("=IFERROR(F{}/D{},0)", row + 1, row + 1);
            ws.write_formula_with_format(row, 6, f.as_str(), &fmt_body_pct)?;
        } else {
            ws.write_blank(row, 6, &fmt_body_pct)?; // G15
        }

        // H: Input (Right Edge -> Medium Border)
        let fmt_h = fmt_inp_txt.clone().set_border_right(border_medium);
        ws.write_blank(row, 7, &fmt_h)?;
    }

    // --- Row 19 (Excel 20): Summary Row ---
    let fmt_sum_lbl = fmt_lbl
        .clone()
        .set_background_color(fill_summary)
        .set_border_bottom(border_medium);

    let fmt_sum_val = fmt_body_num
        .clone()
        .set_background_color(fill_summary)
        .set_border_bottom(border_medium);

    ws.merge_range(19, 1, 19, 2, "", &fmt_sum_lbl)?;
    ws.write_formula_with_format(
        19,
        1,
        "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,21,FALSE))",
        &fmt_sum_lbl,
    )?;

    ws.write_formula_with_format(19, 3, "=SUMPRODUCT(ROUND(D15:D19, 2))", &fmt_sum_val)?;
    ws.write_formula_with_format(19, 4, "=SUMPRODUCT(ROUND(E15:E19, 2))", &fmt_sum_val)?;
    ws.write_formula_with_format(19, 5, "=SUMPRODUCT(ROUND(F15:F19, 2))", &fmt_sum_val)?;

    // G20
    ws.write_formula_with_format(
        19,
        6,
        "=IFERROR(INDEX($F$1:$F$1001,ROW())/INDEX($D$1:$D$1001,ROW()),0)",
        &fmt_sum_val.clone().set_num_format("0%"),
    )?;

    // H20
    let fmt_sum_h = fmt_sum_val.clone().set_border_right(border_medium);
    ws.write_blank(19, 7, &fmt_sum_h)?;

    // --- Spacer Rows 20, 21 ---
    ws.set_row_height(20, 13.5)?;

    // --- Footer Table (Rows 22-25) ---
    // Row 22 (Excel 23)
    let fmt_ft_lbl = Format::new()
        .set_align(FormatAlign::Center)
        .set_align(FormatAlign::VerticalCenter)
        .set_text_wrap()
        .set_border_left(border_medium)
        .set_border_top(border_medium)
        .set_border_right(border_thin);

    let fmt_ft_val = Format::new()
        .set_align(FormatAlign::Center)
        .set_align(FormatAlign::VerticalCenter)
        .set_text_wrap()
        .set_border_top(border_medium)
        .set_border_left(border_thin)
        .set_border_right(border_thin);

    let fmt_ft_right = fmt_ft_val.clone().set_border_right(border_medium);

    ws.merge_range(22, 3, 25, 3, "", &fmt_ft_val)?;
    ws.write_formula_with_format(22, 3, "=D11", &fmt_ft_val)?;

    ws.merge_range(22, 4, 25, 4, "", &fmt_ft_val.clone().set_bold())?;
    ws.write_formula_with_format(
        22,
        4,
        "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,25,FALSE))",
        &fmt_ft_val.clone().set_bold(),
    )?;

    ws.merge_range(22, 5, 25, 5, "", &fmt_ft_val)?;
    ws.write_formula_with_format(
        22,
        5,
        "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,55,FALSE))",
        &fmt_ft_val,
    )?;

    ws.merge_range(22, 6, 25, 6, "", &fmt_ft_val)?;
    ws.write_formula_with_format(
        22,
        6,
        "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,56,FALSE))",
        &fmt_ft_val,
    )?;

    ws.merge_range(22, 7, 25, 7, "", &fmt_ft_right)?;
    ws.write_formula_with_format(22, 7, "=H11", &fmt_ft_right)?;

    // B23, C23
    ws.write_blank(22, 1, &fmt_ft_lbl)?;
    ws.write_blank(
        22,
        2,
        &fmt_ft_lbl
            .clone()
            .set_border_left(FormatBorder::None)
            .set_border_right(border_thin),
    )?;

    // Row 23 (Excel 24)
    // B24:C24 Merged.
    let fmt_ft_mid_base = Format::new()
        .set_align(FormatAlign::Center)
        .set_align(FormatAlign::VerticalCenter)
        .set_border_left(border_medium)
        .set_border_right(border_thin);
    let fmt_ft_mid_bold = fmt_ft_mid_base.clone().set_bold();

    ws.merge_range(23, 1, 23, 2, "", &fmt_ft_mid_bold)?;
    ws.write_formula_with_format(
        23,
        1,
        "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,24,FALSE))",
        &fmt_ft_mid_bold,
    )?;

    // Row 24 (Excel 25)
    // B25:C25 Merged.
    ws.merge_range(24, 1, 24, 2, "", &fmt_ft_mid_base)?;
    ws.write_formula_with_format(24, 1, "=B13", &fmt_ft_mid_base)?;

    // Row 25 (Excel 26)
    // B26, C26 blank.
    ws.write_blank(
        25,
        1,
        &fmt_ft_mid_base.clone().set_border_bottom(border_medium),
    )?;
    ws.write_blank(
        25,
        2,
        &Format::new()
            .set_border_bottom(border_medium)
            .set_border_right(border_thin),
    )?;

    // --- Right Panel (Rows 14-31) ---
    // This iterates from Row 13 to 30 (0-based)

    let fmt_rp_idx = Format::new()
        .set_align(FormatAlign::Right)
        .set_border(border_thin);
    let fmt_rp_txt = Format::new()
        .set_align(FormatAlign::Left)
        .set_border(border_thin);
    let fmt_rp_date = Format::new()
        .set_align(FormatAlign::Center)
        .set_num_format("mm-dd-yy")
        .set_border(border_thin)
        .set_background_color(fill_input);
    let fmt_rp_num = Format::new()
        .set_align(FormatAlign::Right)
        .set_num_format("#,##0.00")
        .set_border(border_thin)
        .set_background_color(fill_input);
    let fmt_rp_calc = Format::new()
        .set_align(FormatAlign::Right)
        .set_num_format("0.0000")
        .set_border(border_thin);

    let f_k = "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,23,FALSE))";

    for i in 0..18 {
        let row = 13 + i as u32;
        let num = i + 1;

        // J, K, L, M, N, O
        ws.write_string_with_format(row, 9, &format!("{}. ", num), &fmt_rp_idx)?;
        ws.write_formula_with_format(row, 10, f_k, &fmt_rp_txt)?; // K
        ws.write_blank(row, 11, &fmt_rp_date)?; // L
        ws.write_blank(row, 12, &fmt_rp_num)?; // M
        ws.write_blank(row, 13, &fmt_rp_num)?; // N

        // O: Ratio N/M
        let f_o = format!("=IF(M{}=\"\",\"\",N{}/M{})", row + 1, row + 1, row + 1);
        ws.write_formula_with_format(row, 14, f_o.as_str(), &fmt_rp_calc)?;

        // Q, R, S, T, U, V
        ws.write_string_with_format(row, 16, &format!("{}. ", num + 18), &fmt_rp_idx)?;
        ws.write_formula_with_format(row, 17, f_k, &fmt_rp_txt)?; // R
        ws.write_blank(row, 18, &fmt_rp_date)?; // S
        ws.write_blank(row, 19, &fmt_rp_num)?; // T
        ws.write_blank(row, 20, &fmt_rp_num)?; // U

        let f_v = format!("=IF(T{}=\"\",\"\",U{}/T{})", row + 1, row + 1, row + 1);
        ws.write_formula_with_format(row, 21, f_v.as_str(), &fmt_rp_calc)?;
    }

    Ok(())
}
