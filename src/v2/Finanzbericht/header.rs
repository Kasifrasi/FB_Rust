use super::styles::ReportStyles;
use rust_xlsxwriter::{DataValidation, Format, FormatAlign, FormatBorder, Worksheet, XlsxError};

// ============================================================================
// Datenstrukturen für die getrennte Logik
// ============================================================================

/// Zellposition
struct Cell {
    row: u32,
    col: u16,
}

/// Merge-Bereich Definition
struct MergeRange {
    first_row: u32,
    first_col: u16,
    last_row: u32,
    last_col: u16,
}

/// Formatierte leere Zelle
struct BlankCell<'a> {
    cell: Cell,
    format: &'a Format,
}

/// Formatierte Zelle mit Formel
struct FormulaCell<'a> {
    cell: Cell,
    formula: &'a str,
    format: &'a Format,
}

// ============================================================================
// Lokale Styles
// ============================================================================

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
        let fmt_th_c_top = s.table_header_base.clone().set_border_right(s.border_thin);

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
            .set_border_right(s.border_thin);

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
            .set_border_right(s.border_thin);

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

// ============================================================================
// Hauptfunktion
// ============================================================================

pub fn write_header(
    ws: &mut Worksheet,
    styles: &ReportStyles,
    suffix: &str,
    lang_val: &str,
) -> Result<(), XlsxError> {
    let ls = LocalStyles::new(styles);

    // Durchlauf 1: Formatierung (Merges und leere formatierte Zellen)
    write_formatting(ws, styles, &ls)?;

    // Durchlauf 2: Werte (Strings und Zahlen)
    write_values(ws, styles, suffix, lang_val)?;

    // Durchlauf 3: Formeln
    write_formulas(ws, styles, &ls)?;

    // Durchlauf 4: Table Body
    write_table_body(ws, styles)?;

    Ok(())
}

// ============================================================================
// Durchlauf 1: Formatierung (Merges und Blanks)
// ============================================================================

fn write_formatting(
    ws: &mut Worksheet,
    styles: &ReportStyles,
    ls: &LocalStyles,
) -> Result<(), XlsxError> {
    // Definiere alle Merge-Bereiche mit ihren Formaten
    let merges: Vec<(MergeRange, &Format)> = vec![
        // Row 0
        (
            MergeRange {
                first_row: 0,
                first_col: 1,
                last_row: 0,
                last_col: 2,
            },
            &styles.header_label,
        ),
        // Row 1
        (
            MergeRange {
                first_row: 1,
                first_col: 1,
                last_row: 1,
                last_col: 2,
            },
            &styles.header_suffix,
        ),
        (
            MergeRange {
                first_row: 1,
                first_col: 9,
                last_row: 2,
                last_col: 14,
            },
            &styles.header_right_box_body,
        ),
        // Row 2
        (
            MergeRange {
                first_row: 2,
                first_col: 1,
                last_row: 2,
                last_col: 2,
            },
            &styles.header_label,
        ),
        // Row 3
        (
            MergeRange {
                first_row: 3,
                first_col: 9,
                last_row: 3,
                last_col: 14,
            },
            &styles.link_style,
        ),
        // Row 4
        (
            MergeRange {
                first_row: 4,
                first_col: 1,
                last_row: 4,
                last_col: 2,
            },
            &styles.left_center,
        ),
        // Row 5-6
        (
            MergeRange {
                first_row: 5,
                first_col: 1,
                last_row: 6,
                last_col: 2,
            },
            &styles.left_center,
        ),
        (
            MergeRange {
                first_row: 5,
                first_col: 3,
                last_row: 6,
                last_col: 7,
            },
            &ls.fmt_d6,
        ),
        // Row 7
        (
            MergeRange {
                first_row: 7,
                first_col: 1,
                last_row: 7,
                last_col: 2,
            },
            &styles.left_center,
        ),
        (
            MergeRange {
                first_row: 7,
                first_col: 6,
                last_row: 7,
                last_col: 7,
            },
            &ls.fmt_row7_date,
        ),
        // Row 8
        (
            MergeRange {
                first_row: 8,
                first_col: 1,
                last_row: 8,
                last_col: 2,
            },
            &styles.left_center,
        ),
        (
            MergeRange {
                first_row: 8,
                first_col: 6,
                last_row: 8,
                last_col: 7,
            },
            &ls.fmt_row7_date,
        ),
        // Row 10 (Table Header)
        (
            MergeRange {
                first_row: 10,
                first_col: 9,
                last_row: 10,
                last_col: 10,
            },
            &styles.left_center_bold,
        ),
        (
            MergeRange {
                first_row: 10,
                first_col: 16,
                last_row: 10,
                last_col: 17,
            },
            &styles.left_center_bold,
        ),
        // Vertical Merges D11-H13
        (
            MergeRange {
                first_row: 10,
                first_col: 3,
                last_row: 13,
                last_col: 3,
            },
            &ls.fmt_th_d,
        ),
        (
            MergeRange {
                first_row: 10,
                first_col: 4,
                last_row: 13,
                last_col: 4,
            },
            &ls.fmt_th_d_bold,
        ),
        (
            MergeRange {
                first_row: 10,
                first_col: 5,
                last_row: 13,
                last_col: 5,
            },
            &ls.fmt_th_d,
        ),
        (
            MergeRange {
                first_row: 10,
                first_col: 6,
                last_row: 13,
                last_col: 6,
            },
            &ls.fmt_th_d,
        ),
        (
            MergeRange {
                first_row: 10,
                first_col: 7,
                last_row: 13,
                last_col: 7,
            },
            &ls.fmt_th_h,
        ),
        // Row 11
        (
            MergeRange {
                first_row: 11,
                first_col: 1,
                last_row: 11,
                last_col: 2,
            },
            &ls.fmt_th_side_bold,
        ),
        // Row 12
        (
            MergeRange {
                first_row: 12,
                first_col: 1,
                last_row: 12,
                last_col: 2,
            },
            &ls.fmt_th_side,
        ),
    ];

    // Schreibe alle Merges
    for (range, format) in &merges {
        ws.merge_range(
            range.first_row,
            range.first_col,
            range.last_row,
            range.last_col,
            "",
            format,
        )?;
    }

    // Definiere alle leeren formatierten Zellen
    let blanks: Vec<BlankCell> = vec![
        // Row 0: K1:N1 (Top Border Medium)
        BlankCell {
            cell: Cell { row: 0, col: 10 },
            format: &ls.fmt_top_med,
        },
        BlankCell {
            cell: Cell { row: 0, col: 11 },
            format: &ls.fmt_top_med,
        },
        BlankCell {
            cell: Cell { row: 0, col: 12 },
            format: &ls.fmt_top_med,
        },
        BlankCell {
            cell: Cell { row: 0, col: 13 },
            format: &ls.fmt_top_med,
        },
        // Row 0: O1 (Top Right Corner)
        BlankCell {
            cell: Cell { row: 0, col: 14 },
            format: &ls.fmt_top_right_med,
        },
        // Row 4: D5
        BlankCell {
            cell: Cell { row: 4, col: 3 },
            format: &ls.fmt_d5,
        },
        // Row 6: J7
        BlankCell {
            cell: Cell { row: 6, col: 9 },
            format: &ls.orange_dotted,
        },
        // Row 7: E8, J8
        BlankCell {
            cell: Cell { row: 7, col: 4 },
            format: &ls.fmt_row7_date,
        },
        BlankCell {
            cell: Cell { row: 7, col: 9 },
            format: &ls.value_dotted,
        },
        // Row 8: J9
        BlankCell {
            cell: Cell { row: 8, col: 9 },
            format: &ls.input_dotted,
        },
        // Row 10: B11, C11
        BlankCell {
            cell: Cell { row: 10, col: 1 },
            format: &ls.fmt_th_b,
        },
        BlankCell {
            cell: Cell { row: 10, col: 2 },
            format: &ls.fmt_th_c_top,
        },
        // Row 13: B14, C14
        BlankCell {
            cell: Cell { row: 13, col: 1 },
            format: &ls.fmt_th_bot_side,
        },
        BlankCell {
            cell: Cell { row: 13, col: 2 },
            format: &ls.fmt_th_bot_right,
        },
    ];

    // Schreibe alle Blanks
    for blank in &blanks {
        ws.write_blank(blank.cell.row, blank.cell.col, blank.format)?;
    }

    Ok(())
}

// ============================================================================
// Durchlauf 2: Werte (Strings und Zahlen)
// ============================================================================

fn write_values(
    ws: &mut Worksheet,
    styles: &ReportStyles,
    suffix: &str,
    lang_val: &str,
) -> Result<(), XlsxError> {
    // String-Werte
    ws.write_string_with_format(1, 1, suffix, &styles.header_suffix)?;
    ws.write_string_with_format(1, 4, lang_val, &styles.input_orange_dashed)?;
    ws.write_string_with_format(12, 12, "Euro", &styles.center_center_bold)?;
    ws.write_string_with_format(12, 19, "Euro", &styles.center_center_bold)?;

    // Data Validation für E2
    let validation = DataValidation::new().allow_list_formula("=Sprachversionen!$B$1:$B$5".into());
    ws.add_data_validation(1, 4, 1, 4, &validation)?;

    Ok(())
}

// ============================================================================
// Durchlauf 3: Formeln
// ============================================================================

fn write_formulas(
    ws: &mut Worksheet,
    styles: &ReportStyles,
    ls: &LocalStyles,
) -> Result<(), XlsxError> {
    // Definiere alle Formeln mit ihren Formaten
    let formulas: Vec<FormulaCell> = vec![
        // Row 0
        FormulaCell {
            cell: Cell { row: 0, col: 1 },
            formula: "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,2,FALSE))",
            format: &styles.header_label,
        },
        FormulaCell {
            cell: Cell { row: 0, col: 9 },
            formula: "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,60,FALSE))",
            format: &styles.header_right_box_top,
        },
        // Row 1
        FormulaCell {
            cell: Cell { row: 1, col: 3 },
            formula: "=IF($E$2=\"\",\"Chose your language\",VLOOKUP($E$2,Sprachversionen!$B:$BN,27,FALSE))",
            format: &styles.left_center,
        },
        FormulaCell {
            cell: Cell { row: 1, col: 9 },
            formula: "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,61,FALSE))",
            format: &styles.header_right_box_body,
        },
        // Row 2
        FormulaCell {
            cell: Cell { row: 2, col: 1 },
            formula: "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,3,FALSE))",
            format: &styles.header_label,
        },
        FormulaCell {
            cell: Cell { row: 2, col: 3 },
            formula: "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,28,FALSE))",
            format: &styles.left_center,
        },
        FormulaCell {
            cell: Cell { row: 2, col: 4 },
            formula: "=IF($E$2=\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,28,FALSE),E3)",
            format: &styles.input_orange_dashed,
        },
        // Row 3
        FormulaCell {
            cell: Cell { row: 3, col: 9 },
            formula: "=HYPERLINK(VLOOKUP($E$2,Sprachversionen!$B:$BN,62,FALSE))",
            format: &styles.link_style,
        },
        // Row 4
        FormulaCell {
            cell: Cell { row: 4, col: 1 },
            formula: "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,4,FALSE))",
            format: &styles.left_center,
        },
        // Row 5
        FormulaCell {
            cell: Cell { row: 5, col: 1 },
            formula: "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,26,FALSE))",
            format: &styles.left_center,
        },
        // Row 7
        FormulaCell {
            cell: Cell { row: 7, col: 1 },
            formula: "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,5,FALSE))",
            format: &styles.left_center,
        },
        FormulaCell {
            cell: Cell { row: 7, col: 3 },
            formula: "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,7,FALSE))",
            format: &ls.fmt_row7_base,
        },
        FormulaCell {
            cell: Cell { row: 7, col: 5 },
            formula: "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,8,FALSE))",
            format: &ls.fmt_row7_base,
        },
        // Row 8
        FormulaCell {
            cell: Cell { row: 8, col: 1 },
            formula: "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,6,FALSE))",
            format: &styles.left_center,
        },
        FormulaCell {
            cell: Cell { row: 8, col: 3 },
            formula: "=D8",
            format: &ls.fmt_row7_base,
        },
        FormulaCell {
            cell: Cell { row: 8, col: 4 },
            formula: "",
            format: &ls.fmt_row7_date,
        },
        FormulaCell {
            cell: Cell { row: 8, col: 5 },
            formula: "=F8",
            format: &ls.fmt_row7_base,
        },
        FormulaCell {
            cell: Cell { row: 8, col: 6 },
            formula: "",
            format: &ls.fmt_row7_date,
        },
        // Row 10 (Header Table)
        FormulaCell {
            cell: Cell { row: 10, col: 9 },
            formula: "=B18",
            format: &styles.left_center_bold,
        },
        FormulaCell {
            cell: Cell { row: 10, col: 16 },
            formula: "=B18",
            format: &styles.left_center_bold,
        },
        // Vertical columns D11-H11
        FormulaCell {
            cell: Cell { row: 10, col: 3 },
            formula: "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,11,FALSE))",
            format: &ls.fmt_th_d,
        },
        FormulaCell {
            cell: Cell { row: 10, col: 4 },
            formula: "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,12,FALSE))",
            format: &ls.fmt_th_d_bold,
        },
        FormulaCell {
            cell: Cell { row: 10, col: 5 },
            formula: "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,13,FALSE))",
            format: &ls.fmt_th_d,
        },
        FormulaCell {
            cell: Cell { row: 10, col: 6 },
            formula: "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,14,FALSE))",
            format: &ls.fmt_th_d,
        },
        FormulaCell {
            cell: Cell { row: 10, col: 7 },
            formula: "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,15,FALSE))",
            format: &ls.fmt_th_h,
        },
        // Row 11
        FormulaCell {
            cell: Cell { row: 11, col: 1 },
            formula: "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,9,FALSE))",
            format: &ls.fmt_th_side_bold,
        },
        // Row 12
        FormulaCell {
            cell: Cell { row: 12, col: 1 },
            formula: "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,10,FALSE))",
            format: &ls.fmt_th_side,
        },
        FormulaCell {
            cell: Cell { row: 12, col: 11 },
            formula: "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,22,FALSE))",
            format: &styles.center_center_bold,
        },
        FormulaCell {
            cell: Cell { row: 12, col: 13 },
            formula: "=IF(E3=\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,28,FALSE),E3)",
            format: &styles.center_center_bold,
        },
        FormulaCell {
            cell: Cell { row: 12, col: 14 },
            formula: "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,58,FALSE))",
            format: &styles.center_center_bold,
        },
        FormulaCell {
            cell: Cell { row: 12, col: 18 },
            formula: "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,22,FALSE))",
            format: &styles.center_center_bold,
        },
        FormulaCell {
            cell: Cell { row: 12, col: 20 },
            formula: "=IF(E3=\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,28,FALSE),E3)",
            format: &styles.center_center_bold,
        },
        FormulaCell {
            cell: Cell { row: 12, col: 21 },
            formula: "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,58,FALSE))",
            format: &styles.center_center_bold,
        },
    ];

    // Schreibe alle Formeln
    for fc in &formulas {
        if fc.formula.is_empty() {
            ws.write_blank(fc.cell.row, fc.cell.col, fc.format)?;
        } else {
            ws.write_formula_with_format(fc.cell.row, fc.cell.col, fc.formula, fc.format)?;
        }
    }

    // K8 ohne Format
    ws.write_formula(
        7,
        10,
        "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,59,FALSE))",
    )?;

    Ok(())
}

// ============================================================================
// Durchlauf 4: Table Body
// ============================================================================

fn write_table_body(ws: &mut Worksheet, styles: &ReportStyles) -> Result<(), XlsxError> {
    let vlookup_indices = [16, 17, 18, 19, 20];

    let body_label_no_h = styles
        .body_label
        .clone()
        .set_border_top(FormatBorder::None)
        .set_border_bottom(FormatBorder::None);

    let body_label_top = styles
        .body_label
        .clone()
        .set_border_bottom(FormatBorder::None);

    // --- Durchlauf 4a: Formatierung ---
    for (i, _) in vlookup_indices.iter().enumerate() {
        let row = 14 + i as u32;
        let style = if i == 0 {
            &body_label_top
        } else {
            &body_label_no_h
        };

        ws.merge_range(row, 1, row, 2, "", style)?;
        ws.write_blank(row, 4, &styles.body_input)?;
        ws.write_blank(row, 5, &styles.body_calc)?;
        ws.write_blank(row, 7, &styles.body_right)?;
    }

    // Summary Row Formatierung
    ws.merge_range(19, 1, 19, 2, "", &styles.summary_label)?;
    ws.write_blank(19, 7, &styles.summary_right)?;

    // --- Durchlauf 4b: Werte ---
    for (i, _) in vlookup_indices.iter().enumerate() {
        let row = 14 + i as u32;
        ws.write_number_with_format(row, 3, 0.0, &styles.body_value)?;
    }

    // --- Durchlauf 4c: Formeln ---
    for (i, v_idx) in vlookup_indices.iter().enumerate() {
        let row = 14 + i as u32;
        let style = if i == 0 {
            &body_label_top
        } else {
            &body_label_no_h
        };

        // B:C Label Formula
        let formula_b = format!(
            "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,{},FALSE))",
            v_idx
        );
        ws.write_formula_with_format(row, 1, formula_b.as_str(), style)?;

        // G: % Formula
        if row >= 15 {
            let formula_g = format!("=IFERROR(F{}/D{},0)", row + 1, row + 1);
            ws.write_formula_with_format(row, 6, formula_g.as_str(), &styles.body_pct)?;
        } else {
            ws.write_blank(row, 6, &styles.body_pct)?;
        }
    }

    // Summary Row Formeln
    ws.write_formula_with_format(
        19,
        1,
        "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,21,FALSE))",
        &styles.summary_label,
    )?;
    ws.write_formula_with_format(
        19,
        3,
        "=SUMPRODUCT(ROUND(D15:D19, 2))",
        &styles.summary_value,
    )?;
    ws.write_formula_with_format(
        19,
        4,
        "=SUMPRODUCT(ROUND(E15:E19, 2))",
        &styles.summary_value,
    )?;
    ws.write_formula_with_format(
        19,
        5,
        "=SUMPRODUCT(ROUND(F15:F19, 2))",
        &styles.summary_value,
    )?;
    ws.write_formula_with_format(
        19,
        6,
        "=IFERROR(INDEX($F$1:$F$1001,ROW())/INDEX($D$1:$D$1001,ROW()),0)",
        &styles.summary_pct,
    )?;

    Ok(())
}
