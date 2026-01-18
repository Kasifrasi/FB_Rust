use crate::formats::StylePalette;
use crate::sheet_builder::{addr, hide_column, with_border};
use anyhow::Result;
use rust_xlsxwriter::{
    Color, Format, FormatAlign, FormatBorder, FormatPattern, Formula, Worksheet,
};
use std::collections::HashMap;

const CATEGORY_BODY_START_ROW: u32 = 26; // Excel row 27
const SINGLE_LINE_CATEGORIES: [u8; 3] = [6, 7, 8];
const RATIO_FORMULA: &str = "=IFERROR(INDEX($F$1:$F$1001,ROW())/INDEX($D$1:$D$1001,ROW()),0)";
const SIDE_COLUMNS_TO_HIDE: &[char] = &['Q', 'R', 'S', 'T', 'U', 'V'];

// Mapping category -> Sprachversionen lookup index for labels.
const LABEL_LOOKUP_INDEX: &[(u8, u32)] = &[
    (1, 29),
    (2, 31),
    (3, 33),
    (4, 35),
    (5, 37),
    (6, 39),
    (7, 40),
    (8, 41),
];

fn category_label(cat_num: u8) -> Formula {
    let idx = LABEL_LOOKUP_INDEX
        .iter()
        .find(|(c, _)| *c == cat_num)
        .map(|(_, i)| *i)
        .unwrap_or(29);
    vlookup_formula(idx)
}

fn vlookup_formula(idx: u32) -> Formula {
    Formula::new(format!(
        "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,{idx},FALSE))"
    ))
}

fn bold(base: &Format) -> Format {
    base.clone().set_bold()
}

pub fn build_dynamic_report(
    ws: &mut Worksheet,
    palette: &StylePalette,
    num_positions_by_cat: &HashMap<u8, u16>,
) -> Result<u32> {
    let mut current_row = CATEGORY_BODY_START_ROW;
    let mut sum_rows: Vec<u32> = Vec::new();
    let mut ratio_rows: Vec<u32> = Vec::new();
    let mut extra_totals: HashMap<char, Vec<String>> =
        HashMap::from([('D', Vec::new()), ('E', Vec::new()), ('F', Vec::new())]);

    for cat_num in 1..=8 {
        if SINGLE_LINE_CATEGORIES.contains(&cat_num) {
            write_single_category_row(ws, palette, current_row, cat_num)?;
            ratio_rows.push(current_row);
            for col in ['D', 'E', 'F'] {
                extra_totals
                    .get_mut(&col)
                    .unwrap()
                    .push(addr(col, current_row));
            }
            current_row += 1;
            continue;
        }

        write_category_header(ws, palette, current_row, cat_num)?;
        current_row += 1;

        let num_pos = *num_positions_by_cat.get(&cat_num).unwrap_or(&20);
        let pos_start = current_row;
        for pos_num in 1..=num_pos {
            write_position_row(ws, palette, current_row, cat_num, pos_num)?;
            ratio_rows.push(current_row);
            current_row += 1;
        }
        let pos_end = current_row - 1;

        write_category_footer(ws, palette, current_row, cat_num, pos_start, pos_end)?;
        ratio_rows.push(current_row);
        sum_rows.push(current_row);
        current_row += 1;
    }

    let total_row = current_row;
    write_report_total(ws, palette, total_row, &sum_rows, &extra_totals)?;
    ratio_rows.push(total_row);

    let footer_start = total_row + 4;
    let last_row = insert_footer(ws, palette, footer_start, total_row)?;

    apply_ratio_formulas(ws, palette, &ratio_rows)?;
    hide_side_columns(ws);

    Ok(last_row)
}

fn write_category_header(
    ws: &mut Worksheet,
    palette: &StylePalette,
    row: u32,
    cat_num: u8,
) -> Result<()> {
    let fmt_b = with_border(
        &palette.align_right.clone().set_bold().set_num_format("@"),
        Some(FormatBorder::Medium),
        Some(FormatBorder::Thin),
        Some(FormatBorder::Thin),
        Some(FormatBorder::Thin),
    );
    ws.write_string_with_format(row, 1, format!("{cat_num}."), &fmt_b)?;

    let fmt_c = with_border(
        &palette.align_left.clone().set_bold(),
        Some(FormatBorder::Thin),
        Some(FormatBorder::Thin),
        Some(FormatBorder::Thin),
        Some(FormatBorder::Thin),
    );
    ws.write_formula_with_format(row, 2, category_label(cat_num), &fmt_c)?;

    let fmt_values = with_border(
        &palette
            .align_right_wrap
            .clone()
            .set_bold()
            .set_num_format("#,##0.00"),
        Some(FormatBorder::Thin),
        Some(FormatBorder::Thin),
        Some(FormatBorder::Thin),
        Some(FormatBorder::Thin),
    );
    for col in 3..=5 {
        ws.write_blank(row, col, &fmt_values)?;
    }

    let fmt_percent = with_border(
        &palette
            .align_percent
            .clone()
            .set_bold()
            .set_num_format("0%"),
        Some(FormatBorder::Thin),
        Some(FormatBorder::Thin),
        Some(FormatBorder::Thin),
        Some(FormatBorder::Thin),
    );
    ws.write_blank(row, 6, &fmt_percent)?;

    let fmt_h = with_border(
        &palette.align_left.clone().set_bold(),
        Some(FormatBorder::Thin),
        Some(FormatBorder::Medium),
        Some(FormatBorder::Thin),
        Some(FormatBorder::Thin),
    );
    ws.write_blank(row, 7, &fmt_h)?;

    Ok(())
}

fn write_position_row(
    ws: &mut Worksheet,
    palette: &StylePalette,
    row: u32,
    cat_num: u8,
    pos_num: u16,
) -> Result<()> {
    let fmt_b = with_border(
        &palette.align_right.clone().set_num_format("@"),
        Some(FormatBorder::Medium),
        Some(FormatBorder::Thin),
        Some(FormatBorder::Thin),
        Some(FormatBorder::Thin),
    );
    ws.write_string_with_format(row, 1, format!("{cat_num}.{pos_num}"), &fmt_b)?;

    let fmt_c = with_border(
        &palette.fill_input.clone().set_text_wrap(),
        Some(FormatBorder::Thin),
        Some(FormatBorder::Thin),
        Some(FormatBorder::Thin),
        Some(FormatBorder::Thin),
    );
    ws.write_blank(row, 2, &fmt_c)?;

    let fmt_d = with_border(
        &palette
            .fill_value
            .clone()
            .set_text_wrap()
            .set_num_format("#,##0.00"),
        Some(FormatBorder::Thin),
        Some(FormatBorder::Thin),
        Some(FormatBorder::Thin),
        Some(FormatBorder::Thin),
    );
    ws.write_blank(row, 3, &fmt_d)?;

    let fmt_e = with_border(
        &palette
            .fill_input
            .clone()
            .set_text_wrap()
            .set_num_format("#,##0.00"),
        Some(FormatBorder::Thin),
        Some(FormatBorder::Thin),
        Some(FormatBorder::Thin),
        Some(FormatBorder::Thin),
    );
    ws.write_blank(row, 4, &fmt_e)?;

    let fmt_f = with_border(
        &palette
            .fill_input
            .clone()
            .set_align(FormatAlign::Right)
            .set_num_format("#,##0.00"),
        Some(FormatBorder::Thin),
        Some(FormatBorder::Thin),
        Some(FormatBorder::Thin),
        Some(FormatBorder::Thin),
    );
    ws.write_blank(row, 5, &fmt_f)?;

    let fmt_g = with_border(
        &palette.align_percent.clone().set_num_format("0%"),
        Some(FormatBorder::Thin),
        Some(FormatBorder::Thin),
        Some(FormatBorder::Thin),
        Some(FormatBorder::Thin),
    );
    ws.write_blank(row, 6, &fmt_g)?;

    let fmt_h = with_border(
        &palette.fill_input.clone().set_text_wrap(),
        Some(FormatBorder::Thin),
        Some(FormatBorder::Medium),
        Some(FormatBorder::Thin),
        Some(FormatBorder::Thin),
    );
    ws.write_blank(row, 7, &fmt_h)?;

    Ok(())
}

fn write_single_category_row(
    ws: &mut Worksheet,
    palette: &StylePalette,
    row: u32,
    cat_num: u8,
) -> Result<()> {
    write_position_row(ws, palette, row, cat_num, 1)?;

    let fmt_b = with_border(
        &palette.align_right.clone().set_bold().set_num_format("@"),
        Some(FormatBorder::Medium),
        Some(FormatBorder::Thin),
        Some(FormatBorder::Thin),
        Some(FormatBorder::Thin),
    );
    ws.write_string_with_format(row, 1, format!("{cat_num}."), &fmt_b)?;

    let mut fmt_c = with_border(
        &palette.fill_white.clone().set_bold(),
        Some(FormatBorder::Thin),
        Some(FormatBorder::Thin),
        Some(FormatBorder::Thin),
        Some(FormatBorder::Thin),
    );
    if cat_num == 8 {
        fmt_c = fmt_c.set_font_color(Color::RGB(0xBFBFBF));
        fmt_c = fmt_c.set_border_bottom(FormatBorder::Thin);
    }
    ws.write_formula_with_format(row, 2, category_label(cat_num), &fmt_c)?;

    Ok(())
}

fn write_category_footer(
    ws: &mut Worksheet,
    palette: &StylePalette,
    row: u32,
    cat_num: u8,
    start_row: u32,
    end_row: u32,
) -> Result<()> {
    let base = palette.fill_summary.clone().set_bold();

    let fmt_b = with_border(
        &base,
        Some(FormatBorder::Medium),
        Some(FormatBorder::Thin),
        Some(FormatBorder::Thin),
        Some(FormatBorder::Thin),
    );
    let vlookup_idx = 30 + (cat_num as u32 - 1) * 2;
    ws.write_formula_with_format(row, 1, vlookup_formula(vlookup_idx), &fmt_b)?;

    let fmt_c = with_border(
        &base,
        Some(FormatBorder::Thin),
        Some(FormatBorder::Thin),
        Some(FormatBorder::Thin),
        Some(FormatBorder::Thin),
    );
    ws.write_blank(row, 2, &fmt_c)?;

    let fmt_sum = with_border(
        &base
            .clone()
            .set_align(FormatAlign::Right)
            .set_num_format("#,##0.00"),
        Some(FormatBorder::Thin),
        Some(FormatBorder::Thin),
        Some(FormatBorder::Thin),
        Some(FormatBorder::Thin),
    );
    for col in ['D', 'E', 'F'] {
        let formula_str = format!(
            "=SUMPRODUCT(ROUND({col}{start}:{col}{end},2))",
            start = start_row + 1,
            end = end_row + 1
        );
        let col_idx = col as u8 - b'A';
        let formula = Formula::new(formula_str);
        ws.write_formula_with_format(row, col_idx as u16, formula, &fmt_sum)?;
    }

    // Simply use the base gray format with borders, no number format, no formula
    let fmt_gray_simple = with_border(
        &base,
        Some(FormatBorder::Thin),
        Some(FormatBorder::Thin),
        Some(FormatBorder::Thin),
        Some(FormatBorder::Thin),
    );
    ws.write_string_with_format(row, 6, " ", &fmt_gray_simple)?;

    let fmt_h = with_border(
        &base,
        Some(FormatBorder::Thin),
        Some(FormatBorder::Medium),
        Some(FormatBorder::Thin),
        Some(FormatBorder::Thin),
    );
    ws.write_blank(row, 7, &fmt_h)?;

    Ok(())
}

fn write_report_total(
    ws: &mut Worksheet,
    palette: &StylePalette,
    row: u32,
    sum_rows: &[u32],
    extra_totals: &HashMap<char, Vec<String>>,
) -> Result<()> {
    let base = palette.fill_summary.clone().set_bold();

    let fmt_b = with_border(
        &base.clone(),
        Some(FormatBorder::Medium),
        Some(FormatBorder::Thin),
        Some(FormatBorder::Thin),
        Some(FormatBorder::Medium),
    );
    ws.write_formula_with_format(row, 1, vlookup_formula(42), &fmt_b)?;

    let fmt_c = with_border(
        &base.clone(),
        Some(FormatBorder::Thin),
        Some(FormatBorder::Thin),
        Some(FormatBorder::Thin),
        Some(FormatBorder::Medium),
    );
    ws.write_blank(row, 2, &fmt_c)?;

    let fmt_sum = with_border(
        &base
            .clone()
            .set_align(FormatAlign::Right)
            .set_num_format("#,##0.00"),
        Some(FormatBorder::Thin),
        Some(FormatBorder::Thin),
        Some(FormatBorder::Thin),
        Some(FormatBorder::Medium),
    );
    for col in ['D', 'E', 'F'] {
        let mut refs: Vec<String> = sum_rows.iter().map(|r| addr(col, *r)).collect();
        if let Some(extra) = extra_totals.get(&col) {
            refs.extend(extra.clone());
        }
        let formula = format!("=SUM({})", refs.join("+"));
        let col_idx = col as u8 - b'A';
        ws.write_formula_with_format(row, col_idx as u16, Formula::new(formula), &fmt_sum)?;
    }

    // Simply use the base gray format with borders, no number format, no formula
    let fmt_gray_simple = with_border(
        &base,
        Some(FormatBorder::Thin),
        Some(FormatBorder::Thin),
        Some(FormatBorder::Thin),
        Some(FormatBorder::Medium),
    );
    ws.write_string_with_format(row, 6, " ", &fmt_gray_simple)?;

    let fmt_h = with_border(
        &base.clone(),
        Some(FormatBorder::Thin),
        Some(FormatBorder::Medium),
        Some(FormatBorder::Thin),
        Some(FormatBorder::Medium),
    );
    ws.write_blank(row, 7, &fmt_h)?;

    Ok(())
}

fn insert_footer(
    ws: &mut Worksheet,
    palette: &StylePalette,
    footer_start: u32,
    total_row: u32,
) -> Result<u32> {
    // Helpers for borders
    let thin = Some(FormatBorder::Thin);
    let border_center = with_border(&bold(&palette.align_center), thin, thin, thin, thin);
    let border_center_wrap = with_border(&bold(&palette.align_center_wrap), thin, thin, thin, thin);
    let border_left = with_border(&bold(&palette.align_left), thin, thin, thin, thin);
    let border_check = with_border(&palette.align_center, thin, thin, thin, thin);
    let border_diff = with_border(
        &palette
            .fill_input
            .clone()
            .set_align(FormatAlign::Right)
            .set_num_format("#,##0.00"),
        thin,
        thin,
        thin,
        thin,
    );
    let border_ok = with_border(
        &palette
            .align_center
            .clone()
            .set_font_color(Color::RGB(0xBFBFBF)),
        thin,
        thin,
        thin,
        thin,
    );
    let border_cat_left = with_border(&palette.align_left, thin, thin, thin, thin);

    // Headline
    ws.write_formula_with_format(footer_start, 4, vlookup_formula(44), &border_center)?;

    // Laut Fibu merged B:D
    let fibu_text = vlookup_formula(43);
    ws.merge_range(footer_start + 1, 1, footer_start + 1, 3, "", &border_center)?;
    ws.write_formula_with_format(footer_start + 1, 1, fibu_text, &border_center)?;

    // E merged across two rows
    let e_text = vlookup_formula(44);
    ws.merge_range(
        footer_start,
        4,
        footer_start + 1,
        4,
        "",
        &border_center_wrap,
    )?;
    ws.write_formula_with_format(footer_start, 4, e_text, &border_center_wrap)?;

    // Reference row to B13
    ws.merge_range(
        footer_start + 2,
        1,
        footer_start + 2,
        2,
        "",
        &palette.align_left,
    )?;
    ws.write_blank(footer_start + 2, 3, &palette.align_right)?;
    ws.write_formula_with_format(
        footer_start + 2,
        4,
        Formula::new("=B13"),
        &palette.align_right,
    )?;

    // Difference check
    ws.write_formula_with_format(footer_start + 4, 1, vlookup_formula(45), &border_left)?;
    // Fill gaps with border
    ws.write_blank(footer_start + 4, 2, &border_left)?;

    let check_formula = format!(
        "=IF(ROUND(E{};2)=ROUND(F20-F{};2);\"\u{2713}\";\"\")",
        footer_start + 5,
        total_row + 1
    );
    ws.write_formula_with_format(
        footer_start + 4,
        3,
        Formula::new(check_formula),
        &border_check,
    )?;

    let diff_formula = format!("=E20-E{}", total_row + 1);
    ws.write_formula_with_format(
        footer_start + 4,
        4,
        Formula::new(diff_formula),
        &border_diff,
    )?;

    // OK row and category sums 1-3
    ws.write_formula_with_format(footer_start + 6, 1, vlookup_formula(46), &border_left)?;
    ws.write_blank(footer_start + 6, 2, &border_left)?;
    ws.write_blank(footer_start + 6, 3, &border_left)?;
    let ok_formula = format!(
        "=IF(E{}=SUM(E{}:E{});\"OK\";\"\")",
        footer_start + 5,
        footer_start + 7,
        footer_start + 9
    );
    ws.write_formula_with_format(footer_start + 6, 4, Formula::new(ok_formula), &border_ok)?;

    let vlookup_indices = [47, 48, 49];
    for (i, idx) in vlookup_indices.iter().enumerate() {
        let row = footer_start + 7 + i as u32;
        ws.write_formula_with_format(row, 1, vlookup_formula(*idx), &border_cat_left)?;
        ws.write_blank(row, 2, &border_cat_left)?;
        ws.write_blank(row, 3, &border_cat_left)?;
        ws.write_blank(row, 4, &border_diff)?;
    }

    // Generic placeholders and labels
    for offset in 10..=12 {
        let r = footer_start + offset;
        for c in 1..=4 {
            ws.write_blank(r, c, &palette.align_left)?;
        }
    }

    ws.write_formula_with_format(
        footer_start + 13,
        1,
        vlookup_formula(50),
        &palette.align_left,
    )?;
    ws.write_formula_with_format(
        footer_start + 14,
        1,
        vlookup_formula(54),
        &palette.align_left,
    )?;

    for offset in 15..=18 {
        let r = footer_start + offset;
        for c in 1..=4 {
            ws.write_blank(r, c, &palette.align_left)?;
        }
    }

    ws.write_formula_with_format(
        footer_start + 19,
        1,
        vlookup_formula(51),
        &bold(&palette.align_left),
    )?;
    ws.write_blank(footer_start + 19, 2, &palette.align_left)?;
    ws.write_formula_with_format(
        footer_start + 19,
        3,
        vlookup_formula(52),
        &bold(&palette.align_left),
    )?;
    ws.write_blank(footer_start + 19, 4, &palette.align_left)?;

    ws.write_formula_with_format(
        footer_start + 20,
        3,
        vlookup_formula(53),
        &palette.align_left,
    )?;
    for c in 1..=4 {
        ws.write_blank(footer_start + 20, c, &palette.align_left)?;
    }

    Ok(footer_start + 20)
}

fn apply_ratio_formulas(ws: &mut Worksheet, palette: &StylePalette, rows: &[u32]) -> Result<()> {
    let fmt = with_border(
        &palette.align_percent.clone().set_num_format("0%"),
        Some(FormatBorder::Thin),
        Some(FormatBorder::Thin),
        Some(FormatBorder::Thin),
        Some(FormatBorder::Thin),
    );
    for row in rows {
        ws.write_formula_with_format(*row, 6, Formula::new(RATIO_FORMULA), &fmt)?;
    }
    Ok(())
}

fn hide_side_columns(ws: &mut Worksheet) {
    for col in SIDE_COLUMNS_TO_HIDE {
        hide_column(ws, *col);
    }
}
