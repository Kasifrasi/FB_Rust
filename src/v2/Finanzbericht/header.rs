use super::styles::ReportStyles;
use crate::v2::Sprachversion::data::CURRENCIES;
use rust_xlsxwriter::{DataValidation, Format, FormatAlign, FormatBorder, Worksheet, XlsxError};
use std::collections::HashMap;

// ============================================================================
// Format-Matrix: Zentrale Speicherung aller Zellformate
// ============================================================================

/// Zentrale Matrix für alle Zellformate
/// Default-Format ist unlocked (Arial 10), nur Formel-Zellen werden locked
struct FormatMatrix {
    formats: HashMap<(u32, u16), Format>,
}

impl FormatMatrix {
    fn new() -> Self {
        Self {
            formats: HashMap::new(),
        }
    }

    /// Setzt das Format für eine Zelle
    fn set(&mut self, row: u32, col: u16, format: &Format) {
        self.formats.insert((row, col), format.clone());
    }

    /// Holt das Format für eine Zelle (Standard - unlocked durch Default-Format)
    fn get(&self, row: u32, col: u16) -> Option<&Format> {
        self.formats.get(&(row, col))
    }

    /// Holt das Format für eine Zelle mit locked flag (für Formeln)
    fn get_locked(&self, row: u32, col: u16) -> Option<Format> {
        self.formats
            .get(&(row, col))
            .map(|f| f.clone().set_locked())
    }
}

// ============================================================================
// Merge-Definition
// ============================================================================

struct MergeRange {
    first_row: u32,
    first_col: u16,
    last_row: u32,
    last_col: u16,
}

// ============================================================================
// Lokale Styles (abgeleitet von ReportStyles)
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
    body_label_top: Format,
    body_label_no_h: Format,
    // Right Panel Header (Row 11-13)
    rp_hdr_left: Format,      // J12-J13: nur links border
    rp_hdr_mid: Format,       // K-N: keine border
    rp_hdr_mid_top: Format,   // L-N Row 11: nur oben border
    rp_hdr_right: Format,     // O12-O13: nur rechts border
    rp_hdr_right_top: Format, // O11: rechts + oben border
    // Right Panel Body (Row 14-31) - J/K Spalten (kein innerer border)
    rp_idx: Format,      // J14-J30: nur links border
    rp_idx_last: Format, // J31: links + unten border
    rp_txt: Format,      // K14-K30: nur rechts border
    rp_txt_last: Format, // K31: rechts + unten border
    // Right Panel Body (Row 14-31) - L-O Spalten (voller border)
    rp_date: Format,      // L14-L30: border ohne unten
    rp_date_last: Format, // L31: voller border
    rp_num: Format,       // M/N 14-30: border ohne unten
    rp_num_last: Format,  // M/N 31: voller border
    rp_calc: Format,      // O14-O30: border ohne unten
    rp_calc_last: Format, // O31: voller border
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

        let body_label_top = s.body_label.clone().set_border_bottom(FormatBorder::None);

        let body_label_no_h = s
            .body_label
            .clone()
            .set_border_top(FormatBorder::None)
            .set_border_bottom(FormatBorder::None);

        // === Right Panel Header Styles (Row 11-13) ===
        let rp_hdr_left = s.base.clone().set_border_left(s.border_thin);
        let rp_hdr_mid = s.base.clone();
        let rp_hdr_mid_top = rp_hdr_mid.clone().set_border_top(s.border_thin);
        let rp_hdr_right = s.base.clone().set_border_right(s.border_thin);
        let rp_hdr_right_top = rp_hdr_right.clone().set_border_top(s.border_thin);

        // === Right Panel Body Styles (Row 14-31) ===
        // J/K Spalten: kein innerer border
        let rp_idx = s.rp_index.clone().set_border_left(s.border_thin);
        let rp_idx_last = rp_idx.clone().set_border_bottom(s.border_thin);
        let rp_txt = s.rp_text.clone().set_border_right(s.border_thin);
        let rp_txt_last = rp_txt.clone().set_border_bottom(s.border_thin);

        // L-O Spalten: voller border (aber ohne bottom außer in letzter Zeile)
        let rp_date = s
            .rp_date
            .clone()
            .set_border_left(s.border_thin)
            .set_border_right(s.border_thin)
            .set_border_top(s.border_thin);
        let rp_date_last = rp_date.clone().set_border_bottom(s.border_thin);

        let rp_num = s
            .rp_number
            .clone()
            .set_border_left(s.border_thin)
            .set_border_right(s.border_thin)
            .set_border_top(s.border_thin);
        let rp_num_last = rp_num.clone().set_border_bottom(s.border_thin);

        let rp_calc = s
            .rp_calc
            .clone()
            .set_border_left(s.border_thin)
            .set_border_right(s.border_thin)
            .set_border_top(s.border_thin);
        let rp_calc_last = rp_calc.clone().set_border_bottom(s.border_thin);

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
            body_label_top,
            body_label_no_h,
            rp_hdr_left,
            rp_hdr_mid,
            rp_hdr_mid_top,
            rp_hdr_right,
            rp_hdr_right_top,
            rp_idx,
            rp_idx_last,
            rp_txt,
            rp_txt_last,
            rp_date,
            rp_date_last,
            rp_num,
            rp_num_last,
            rp_calc,
            rp_calc_last,
        }
    }
}

// ============================================================================
// Format-Matrix Builder: Zentrale Definition aller Zellformate
// ============================================================================

fn build_format_matrix(styles: &ReportStyles, ls: &LocalStyles) -> FormatMatrix {
    let mut matrix = FormatMatrix::new();

    // ========================================================================
    // HEADER BEREICH (Rows 0-13)
    // ========================================================================

    // --- Row 0 ---
    matrix.set(0, 1, &styles.header_label); // B1 (merge B1:C1)
    matrix.set(0, 9, &styles.header_right_box_top); // J1
    matrix.set(0, 10, &ls.fmt_top_med); // K1
    matrix.set(0, 11, &ls.fmt_top_med); // L1
    matrix.set(0, 12, &ls.fmt_top_med); // M1
    matrix.set(0, 13, &ls.fmt_top_med); // N1
    matrix.set(0, 14, &ls.fmt_top_right_med); // O1

    // --- Row 1 ---
    matrix.set(1, 1, &styles.header_suffix); // B2 (merge B2:C2)
    matrix.set(1, 3, &styles.left_center); // D2
    matrix.set(1, 4, &styles.input_orange_dashed); // E2
    matrix.set(1, 9, &styles.header_right_box_body); // J2 (merge J2:O3)

    // --- Row 2 ---
    matrix.set(2, 1, &styles.header_label); // B3 (merge B3:C3)
    matrix.set(2, 3, &styles.left_center); // D3
    matrix.set(2, 4, &styles.input_orange_dashed); // E3

    // --- Row 3 ---
    matrix.set(3, 9, &styles.link_style); // J4 (merge J4:O4)

    // --- Row 4 ---
    matrix.set(4, 1, &styles.left_center); // B5 (merge B5:C5)
    matrix.set(4, 3, &ls.fmt_d5); // D5

    // --- Row 5 ---
    matrix.set(5, 1, &styles.left_center); // B6 (merge B6:C7)
    matrix.set(5, 3, &ls.fmt_d6); // D6 (merge D6:H7)

    // --- Row 6 ---
    matrix.set(6, 9, &ls.orange_dotted); // J7

    // --- Row 7 ---
    matrix.set(7, 1, &styles.left_center); // B8 (merge B8:C8)
    matrix.set(7, 3, &ls.fmt_row7_base); // D8
    matrix.set(7, 4, &ls.fmt_row7_date); // E8
    matrix.set(7, 5, &ls.fmt_row7_base); // F8
    matrix.set(7, 6, &ls.fmt_row7_date); // G8 (merge G8:H8)
    matrix.set(7, 9, &ls.value_dotted); // J8

    // --- Row 8 ---
    matrix.set(8, 1, &styles.left_center); // B9 (merge B9:C9)
    matrix.set(8, 3, &ls.fmt_row7_base); // D9
    matrix.set(8, 4, &ls.fmt_row7_date); // E9
    matrix.set(8, 5, &ls.fmt_row7_base); // F9
    matrix.set(8, 6, &ls.fmt_row7_date); // G9 (merge G9:H9)
    matrix.set(8, 9, &ls.input_dotted); // J9

    // --- Row 10 (Table Header) ---
    matrix.set(10, 1, &ls.fmt_th_b); // B11
    matrix.set(10, 2, &ls.fmt_th_c_top); // C11
    matrix.set(10, 3, &ls.fmt_th_d); // D11 (merge D11:D14)
    matrix.set(10, 4, &ls.fmt_th_d_bold); // E11 (merge E11:E14)
    matrix.set(10, 5, &ls.fmt_th_d); // F11 (merge F11:F14)
    matrix.set(10, 6, &ls.fmt_th_d); // G11 (merge G11:G14)
    matrix.set(10, 7, &ls.fmt_th_h); // H11 (merge H11:H14)
                                     // J11 und Q11 werden im Right Panel Abschnitt gesetzt

    // --- Row 11 ---
    matrix.set(11, 1, &ls.fmt_th_side_bold); // B12 (merge B12:C12)

    // --- Row 12 ---
    matrix.set(12, 1, &ls.fmt_th_side); // B13 (merge B13:C13)
    matrix.set(12, 11, &styles.center_center_bold); // L13
    matrix.set(12, 12, &styles.center_center_bold); // M13
    matrix.set(12, 13, &styles.center_center_bold); // N13
    matrix.set(12, 14, &styles.center_center_bold); // O13
    matrix.set(12, 18, &styles.center_center_bold); // S13
    matrix.set(12, 19, &styles.center_center_bold); // T13
    matrix.set(12, 20, &styles.center_center_bold); // U13
    matrix.set(12, 21, &styles.center_center_bold); // V13

    // --- Row 13 ---
    matrix.set(13, 1, &ls.fmt_th_bot_side); // B14
    matrix.set(13, 2, &ls.fmt_th_bot_right); // C14

    // ========================================================================
    // TABLE BODY (Rows 14-18)
    // ========================================================================

    // Row 14 (erste Datenzeile - spezielle Border)
    matrix.set(14, 1, &ls.body_label_top); // B15 (merge B15:C15)
    matrix.set(14, 3, &styles.body_value); // D15
    matrix.set(14, 4, &styles.body_input); // E15
    matrix.set(14, 5, &styles.body_calc); // F15
    matrix.set(14, 6, &styles.body_pct); // G15
    matrix.set(14, 7, &styles.body_right); // H15

    // Rows 15-18 (normale Datenzeilen)
    for row in 15..=18 {
        matrix.set(row, 1, &ls.body_label_no_h); // B (merge B:C)
        matrix.set(row, 3, &styles.body_value); // D
        matrix.set(row, 4, &styles.body_input); // E
        matrix.set(row, 5, &styles.body_calc); // F
        matrix.set(row, 6, &styles.body_pct); // G
        matrix.set(row, 7, &styles.body_right); // H
    }

    // ========================================================================
    // SUMMARY ROW (Row 19)
    // ========================================================================

    matrix.set(19, 1, &styles.summary_label); // B20 (merge B20:C20)
    matrix.set(19, 3, &styles.summary_value); // D20
    matrix.set(19, 4, &styles.summary_value); // E20
    matrix.set(19, 5, &styles.summary_value); // F20
    matrix.set(19, 6, &styles.summary_pct); // G20
    matrix.set(19, 7, &styles.summary_right); // H20

    // ========================================================================
    // RIGHT PANEL (Row 10-30, Cols J-O und Q-V)
    // ========================================================================

    // --- Row 10 (J11-O11, Q11-V11): Header mit top border ---
    // J11:K11 und Q11:R11 sind merged und haben bold format
    let j11_fmt = styles
        .left_center_bold
        .clone()
        .set_border_left(styles.border_thin)
        .set_border_top(styles.border_thin);
    let q11_fmt = styles
        .left_center_bold
        .clone()
        .set_border_left(styles.border_thin)
        .set_border_top(styles.border_thin);
    // Linke Seite
    matrix.set(10, 9, &j11_fmt); // J11 (merge J11:K11)
    matrix.set(10, 11, &ls.rp_hdr_mid_top); // L11
    matrix.set(10, 12, &ls.rp_hdr_mid_top); // M11
    matrix.set(10, 13, &ls.rp_hdr_mid_top); // N11
    matrix.set(10, 14, &ls.rp_hdr_right_top); // O11
                                              // Rechte Seite
    matrix.set(10, 16, &q11_fmt); // Q11 (merge Q11:R11)
    matrix.set(10, 18, &ls.rp_hdr_mid_top); // S11
    matrix.set(10, 19, &ls.rp_hdr_mid_top); // T11
    matrix.set(10, 20, &ls.rp_hdr_mid_top); // U11
    matrix.set(10, 21, &ls.rp_hdr_right_top); // V11

    // --- Row 11-12 (J12-O13, Q12-V13): Header ohne top border ---
    for row in 11..=12 {
        // Linke Seite
        matrix.set(row, 9, &ls.rp_hdr_left); // J
        matrix.set(row, 10, &ls.rp_hdr_mid); // K
        matrix.set(row, 11, &ls.rp_hdr_mid); // L
        matrix.set(row, 12, &ls.rp_hdr_mid); // M
        matrix.set(row, 13, &ls.rp_hdr_mid); // N
        matrix.set(row, 14, &ls.rp_hdr_right); // O
                                               // Rechte Seite
        matrix.set(row, 16, &ls.rp_hdr_left); // Q
        matrix.set(row, 17, &ls.rp_hdr_mid); // R
        matrix.set(row, 18, &ls.rp_hdr_mid); // S
        matrix.set(row, 19, &ls.rp_hdr_mid); // T
        matrix.set(row, 20, &ls.rp_hdr_mid); // U
        matrix.set(row, 21, &ls.rp_hdr_right); // V
    }

    // --- Row 13-29 (J14-O30, Q14-V30): Body ohne bottom border ---
    for row in 13..=29 {
        // Linke Seite
        matrix.set(row, 9, &ls.rp_idx); // J: Index
        matrix.set(row, 10, &ls.rp_txt); // K: Text (Formel)
        matrix.set(row, 11, &ls.rp_date); // L: Date
        matrix.set(row, 12, &ls.rp_num); // M: Number
        matrix.set(row, 13, &ls.rp_num); // N: Number
        matrix.set(row, 14, &ls.rp_calc); // O: Calc (Formel)
                                          // Rechte Seite
        matrix.set(row, 16, &ls.rp_idx); // Q: Index
        matrix.set(row, 17, &ls.rp_txt); // R: Text (Formel)
        matrix.set(row, 18, &ls.rp_date); // S: Date
        matrix.set(row, 19, &ls.rp_num); // T: Number
        matrix.set(row, 20, &ls.rp_num); // U: Number
        matrix.set(row, 21, &ls.rp_calc); // V: Calc (Formel)
    }

    // --- Row 30 (J31-O31, Q31-V31): Body mit bottom border ---
    // Linke Seite
    matrix.set(30, 9, &ls.rp_idx_last); // J31
    matrix.set(30, 10, &ls.rp_txt_last); // K31
    matrix.set(30, 11, &ls.rp_date_last); // L31
    matrix.set(30, 12, &ls.rp_num_last); // M31
    matrix.set(30, 13, &ls.rp_num_last); // N31
    matrix.set(30, 14, &ls.rp_calc_last); // O31
                                          // Rechte Seite
    matrix.set(30, 16, &ls.rp_idx_last); // Q31
    matrix.set(30, 17, &ls.rp_txt_last); // R31
    matrix.set(30, 18, &ls.rp_date_last); // S31
    matrix.set(30, 19, &ls.rp_num_last); // T31
    matrix.set(30, 20, &ls.rp_num_last); // U31
    matrix.set(30, 21, &ls.rp_calc_last); // V31

    matrix
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
    let fmt = build_format_matrix(styles, &ls);

    // Durchlauf 1: Formatierung (Merges und leere Zellen)
    write_formatting(ws, &fmt)?;

    // Durchlauf 2: Werte
    write_values(ws, &fmt, suffix, lang_val)?;

    // Durchlauf 3: Formeln (mit locked Format)
    write_formulas(ws, &fmt)?;

    // Durchlauf 4: Right Panel (J11:V31)
    write_right_panel(ws, &fmt)?;

    // Freeze Pane: Zeilen 1-9 immer sichtbar (freeze unter Zeile 9 = Row 9)
    ws.set_freeze_panes(9, 0)?;

    Ok(())
}

// ============================================================================
// Durchlauf 1: Formatierung (Merges und Blanks)
// ============================================================================

fn write_formatting(ws: &mut Worksheet, fmt: &FormatMatrix) -> Result<(), XlsxError> {
    // Alle Merge-Bereiche
    let merges: Vec<(MergeRange, u32, u16)> = vec![
        // (MergeRange, format_row, format_col) - Format wird aus Matrix geholt
        (
            MergeRange {
                first_row: 0,
                first_col: 1,
                last_row: 0,
                last_col: 2,
            },
            0,
            1,
        ),
        (
            MergeRange {
                first_row: 1,
                first_col: 1,
                last_row: 1,
                last_col: 2,
            },
            1,
            1,
        ),
        (
            MergeRange {
                first_row: 1,
                first_col: 9,
                last_row: 2,
                last_col: 14,
            },
            1,
            9,
        ),
        (
            MergeRange {
                first_row: 2,
                first_col: 1,
                last_row: 2,
                last_col: 2,
            },
            2,
            1,
        ),
        (
            MergeRange {
                first_row: 3,
                first_col: 9,
                last_row: 3,
                last_col: 14,
            },
            3,
            9,
        ),
        (
            MergeRange {
                first_row: 4,
                first_col: 1,
                last_row: 4,
                last_col: 2,
            },
            4,
            1,
        ),
        (
            MergeRange {
                first_row: 5,
                first_col: 1,
                last_row: 6,
                last_col: 2,
            },
            5,
            1,
        ),
        (
            MergeRange {
                first_row: 5,
                first_col: 3,
                last_row: 6,
                last_col: 7,
            },
            5,
            3,
        ),
        (
            MergeRange {
                first_row: 7,
                first_col: 1,
                last_row: 7,
                last_col: 2,
            },
            7,
            1,
        ),
        (
            MergeRange {
                first_row: 7,
                first_col: 6,
                last_row: 7,
                last_col: 7,
            },
            7,
            6,
        ),
        (
            MergeRange {
                first_row: 8,
                first_col: 1,
                last_row: 8,
                last_col: 2,
            },
            8,
            1,
        ),
        (
            MergeRange {
                first_row: 8,
                first_col: 6,
                last_row: 8,
                last_col: 7,
            },
            8,
            6,
        ),
        (
            MergeRange {
                first_row: 10,
                first_col: 9,
                last_row: 10,
                last_col: 10,
            },
            10,
            9,
        ),
        (
            MergeRange {
                first_row: 10,
                first_col: 16,
                last_row: 10,
                last_col: 17,
            },
            10,
            16,
        ),
        (
            MergeRange {
                first_row: 10,
                first_col: 3,
                last_row: 13,
                last_col: 3,
            },
            10,
            3,
        ),
        (
            MergeRange {
                first_row: 10,
                first_col: 4,
                last_row: 13,
                last_col: 4,
            },
            10,
            4,
        ),
        (
            MergeRange {
                first_row: 10,
                first_col: 5,
                last_row: 13,
                last_col: 5,
            },
            10,
            5,
        ),
        (
            MergeRange {
                first_row: 10,
                first_col: 6,
                last_row: 13,
                last_col: 6,
            },
            10,
            6,
        ),
        (
            MergeRange {
                first_row: 10,
                first_col: 7,
                last_row: 13,
                last_col: 7,
            },
            10,
            7,
        ),
        (
            MergeRange {
                first_row: 11,
                first_col: 1,
                last_row: 11,
                last_col: 2,
            },
            11,
            1,
        ),
        (
            MergeRange {
                first_row: 12,
                first_col: 1,
                last_row: 12,
                last_col: 2,
            },
            12,
            1,
        ),
        // Table Body Merges
        (
            MergeRange {
                first_row: 14,
                first_col: 1,
                last_row: 14,
                last_col: 2,
            },
            14,
            1,
        ),
        (
            MergeRange {
                first_row: 15,
                first_col: 1,
                last_row: 15,
                last_col: 2,
            },
            15,
            1,
        ),
        (
            MergeRange {
                first_row: 16,
                first_col: 1,
                last_row: 16,
                last_col: 2,
            },
            16,
            1,
        ),
        (
            MergeRange {
                first_row: 17,
                first_col: 1,
                last_row: 17,
                last_col: 2,
            },
            17,
            1,
        ),
        (
            MergeRange {
                first_row: 18,
                first_col: 1,
                last_row: 18,
                last_col: 2,
            },
            18,
            1,
        ),
        // Summary Row Merge
        (
            MergeRange {
                first_row: 19,
                first_col: 1,
                last_row: 19,
                last_col: 2,
            },
            19,
            1,
        ),
    ];

    for (range, fmt_row, fmt_col) in &merges {
        if let Some(format) = fmt.get(*fmt_row, *fmt_col) {
            ws.merge_range(
                range.first_row,
                range.first_col,
                range.last_row,
                range.last_col,
                "",
                format,
            )?;
        }
    }

    // Leere formatierte Zellen (nur Format, kein Inhalt)
    let blank_cells: Vec<(u32, u16)> = vec![
        // Row 0
        (0, 10),
        (0, 11),
        (0, 12),
        (0, 13),
        (0, 14),
        // Row 4
        (4, 3),
        // Row 6
        (6, 9),
        // Row 7
        (7, 4),
        (7, 9),
        // Row 8
        (8, 4),
        (8, 6),
        (8, 9),
        // Row 10
        (10, 1),
        (10, 2),
        // Row 13
        (13, 1),
        (13, 2),
        // Table Body - Input/Calc/Pct/Right columns
        (14, 4),
        (14, 5),
        (14, 6),
        (14, 7),
        (15, 4),
        (15, 5),
        (15, 7),
        (16, 4),
        (16, 5),
        (16, 7),
        (17, 4),
        (17, 5),
        (17, 7),
        (18, 4),
        (18, 5),
        (18, 7),
        // Summary Row
        (19, 7),
    ];

    for (row, col) in &blank_cells {
        if let Some(format) = fmt.get(*row, *col) {
            ws.write_blank(*row, *col, format)?;
        }
    }

    Ok(())
}

// ============================================================================
// Durchlauf 2: Werte
// ============================================================================

fn write_values(
    ws: &mut Worksheet,
    fmt: &FormatMatrix,
    suffix: &str,
    lang_val: &str,
) -> Result<(), XlsxError> {
    // String-Werte: (row, col, value)
    let string_values: Vec<(u32, u16, &str)> = vec![
        (1, 1, suffix),
        (1, 4, lang_val),
        (12, 12, "Euro"),
        (12, 19, "Euro"),
    ];

    for (row, col, value) in &string_values {
        if let Some(format) = fmt.get(*row, *col) {
            ws.write_string_with_format(*row, *col, *value, format)?;
        }
    }

    // Zahlen-Werte: (row, col, value)
    let number_values: Vec<(u32, u16, f64)> = vec![
        (14, 3, 0.0),
        (15, 3, 0.0),
        (16, 3, 0.0),
        (17, 3, 0.0),
        (18, 3, 0.0),
    ];

    for (row, col, value) in &number_values {
        if let Some(format) = fmt.get(*row, *col) {
            ws.write_number_with_format(*row, *col, *value, format)?;
        }
    }

    // Data Validation für E2 (Sprache)
    let validation = DataValidation::new().allow_list_formula("=Sprachversionen!$B$1:$B$5".into());
    ws.add_data_validation(1, 4, 1, 4, &validation)?;

    // Data Validation für E3 (Währung) - dynamisch basierend auf CURRENCIES Vector
    let currency_count = CURRENCIES.len();
    let currency_formula = format!("=Sprachversionen!$A$1:$A${}", currency_count);
    let currency_validation =
        DataValidation::new().allow_list_formula(currency_formula.as_str().into());
    ws.add_data_validation(2, 4, 2, 4, &currency_validation)?;

    Ok(())
}

// ============================================================================
// Durchlauf 3: Formeln (mit locked Format)
// ============================================================================

fn write_formulas(ws: &mut Worksheet, fmt: &FormatMatrix) -> Result<(), XlsxError> {
    // Alle Formeln: (row, col, formula) - werden mit locked Format geschrieben
    let formulas: Vec<(u32, u16, &str)> = vec![
        // Row 0
        (
            0,
            1,
            "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,2,FALSE))",
        ),
        (
            0,
            9,
            "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,60,FALSE))",
        ),
        // Row 1
        (
            1,
            3,
            "=IF($E$2=\"\",\"Chose your language\",VLOOKUP($E$2,Sprachversionen!$B:$BN,27,FALSE))",
        ),
        (
            1,
            9,
            "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,61,FALSE))",
        ),
        // Row 2
        (
            2,
            1,
            "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,3,FALSE))",
        ),
        (
            2,
            3,
            "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,28,FALSE))",
        ),
        // E3 (2, 4) - keine Formel mehr, stattdessen Currency-Validierung in write_values()
        // Row 3
        (
            3,
            9,
            "=HYPERLINK(VLOOKUP($E$2,Sprachversionen!$B:$BN,62,FALSE))",
        ),
        // Row 4
        (
            4,
            1,
            "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,4,FALSE))",
        ),
        // Row 5
        (
            5,
            1,
            "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,26,FALSE))",
        ),
        // Row 7
        (
            7,
            1,
            "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,5,FALSE))",
        ),
        (
            7,
            3,
            "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,7,FALSE))",
        ),
        (
            7,
            5,
            "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,8,FALSE))",
        ),
        // Row 8
        (
            8,
            1,
            "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,6,FALSE))",
        ),
        // D9 und F9 zeigen auf D8/F8 - verwende die gleiche VLOOKUP Formel
        (
            8,
            3,
            "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,7,FALSE))",
        ),
        (
            8,
            5,
            "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,8,FALSE))",
        ),
        // Row 10 - J11 und Q11 mit VLOOKUP Formel (zeigt auf B18 Formel-Inhalt)
        (
            10,
            9,
            "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,19,FALSE))",
        ),
        (
            10,
            16,
            "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,19,FALSE))",
        ),
        (
            10,
            3,
            "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,11,FALSE))",
        ),
        (
            10,
            4,
            "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,12,FALSE))",
        ),
        (
            10,
            5,
            "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,13,FALSE))",
        ),
        (
            10,
            6,
            "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,14,FALSE))",
        ),
        (
            10,
            7,
            "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,15,FALSE))",
        ),
        // Row 11
        (
            11,
            1,
            "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,9,FALSE))",
        ),
        // Row 12
        (
            12,
            1,
            "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,10,FALSE))",
        ),
        (
            12,
            11,
            "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,22,FALSE))",
        ),
        (
            12,
            13,
            "=IF(E3=\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,28,FALSE),E3)",
        ),
        (
            12,
            14,
            "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,58,FALSE))",
        ),
        (
            12,
            18,
            "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,22,FALSE))",
        ),
        (
            12,
            20,
            "=IF(E3=\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,28,FALSE),E3)",
        ),
        (
            12,
            21,
            "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,58,FALSE))",
        ),
        // Table Body Labels (VLOOKUP indices 16-20)
        (
            14,
            1,
            "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,16,FALSE))",
        ),
        (
            15,
            1,
            "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,17,FALSE))",
        ),
        (
            16,
            1,
            "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,18,FALSE))",
        ),
        (
            17,
            1,
            "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,19,FALSE))",
        ),
        (
            18,
            1,
            "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,20,FALSE))",
        ),
        // Table Body % Formulas (rows 15-18, not row 14)
        (15, 6, "=IFERROR(F16/D16,0)"),
        (16, 6, "=IFERROR(F17/D17,0)"),
        (17, 6, "=IFERROR(F18/D18,0)"),
        (18, 6, "=IFERROR(F19/D19,0)"),
        // Summary Row
        (
            19,
            1,
            "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,21,FALSE))",
        ),
        (19, 3, "=SUMPRODUCT(ROUND(D15:D19, 2))"),
        (19, 4, "=SUMPRODUCT(ROUND(E15:E19, 2))"),
        (19, 5, "=SUMPRODUCT(ROUND(F15:F19, 2))"),
        (
            19,
            6,
            "=IFERROR(INDEX($F$1:$F$1001,ROW())/INDEX($D$1:$D$1001,ROW()),0)",
        ),
    ];

    for (row, col, formula) in &formulas {
        // Formel-Zellen werden mit locked Format geschrieben
        if let Some(format) = fmt.get_locked(*row, *col) {
            ws.write_formula_with_format(*row, *col, *formula, &format)?;
        }
    }

    // K8 ohne Format (spezialfall) - wird auch locked
    let locked_format = Format::new().set_locked();
    ws.write_formula_with_format(
        7,
        10,
        "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,59,FALSE))",
        &locked_format,
    )?;

    Ok(())
}

// ============================================================================
// Durchlauf 4: Right Panel (J11:V31) - Werte und Formeln
// ============================================================================

fn write_right_panel(ws: &mut Worksheet, fmt: &FormatMatrix) -> Result<(), XlsxError> {
    let f_k = "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,23,FALSE))";

    // Right Panel Header (Row 11-13): Blanks schreiben
    // Row 11 hat J11:K11 und Q11:R11 merges mit Formeln (werden in write_formulas geschrieben)
    // L11-O11 und S11-V11 sind blanks
    for col in [11u16, 12, 13, 14] {
        if let Some(format) = fmt.get(10, col) {
            ws.write_blank(10, col, format)?;
        }
    }
    for col in [18u16, 19, 20, 21] {
        if let Some(format) = fmt.get(10, col) {
            ws.write_blank(10, col, format)?;
        }
    }

    // Row 12-13: Alle Zellen sind blanks außer L13-O13 und S13-V13 (Formeln in write_formulas)
    for row in 11..=12u32 {
        for col in [9u16, 10, 11, 12, 13, 14] {
            if let Some(format) = fmt.get(row, col) {
                ws.write_blank(row, col, format)?;
            }
        }
        for col in [16u16, 17, 18, 19, 20, 21] {
            if let Some(format) = fmt.get(row, col) {
                ws.write_blank(row, col, format)?;
            }
        }
    }

    // Right Panel Body (Row 14-31): 18 Zeilen statisch
    // J: Index (String), K: Text (Formel), L: Date (blank), M/N: Number (blank), O: Calc (Formel)
    // Q: Index (String), R: Text (Formel), S: Date (blank), T/U: Number (blank), V: Calc (Formel)

    // Zeile 14 (Row 13)
    write_rp_row(ws, fmt, 13, 1, 19, f_k)?;
    // Zeile 15 (Row 14)
    write_rp_row(ws, fmt, 14, 2, 20, f_k)?;
    // Zeile 16 (Row 15)
    write_rp_row(ws, fmt, 15, 3, 21, f_k)?;
    // Zeile 17 (Row 16)
    write_rp_row(ws, fmt, 16, 4, 22, f_k)?;
    // Zeile 18 (Row 17)
    write_rp_row(ws, fmt, 17, 5, 23, f_k)?;
    // Zeile 19 (Row 18)
    write_rp_row(ws, fmt, 18, 6, 24, f_k)?;
    // Zeile 20 (Row 19)
    write_rp_row(ws, fmt, 19, 7, 25, f_k)?;
    // Zeile 21 (Row 20)
    write_rp_row(ws, fmt, 20, 8, 26, f_k)?;
    // Zeile 22 (Row 21)
    write_rp_row(ws, fmt, 21, 9, 27, f_k)?;
    // Zeile 23 (Row 22)
    write_rp_row(ws, fmt, 22, 10, 28, f_k)?;
    // Zeile 24 (Row 23)
    write_rp_row(ws, fmt, 23, 11, 29, f_k)?;
    // Zeile 25 (Row 24)
    write_rp_row(ws, fmt, 24, 12, 30, f_k)?;
    // Zeile 26 (Row 25)
    write_rp_row(ws, fmt, 25, 13, 31, f_k)?;
    // Zeile 27 (Row 26)
    write_rp_row(ws, fmt, 26, 14, 32, f_k)?;
    // Zeile 28 (Row 27)
    write_rp_row(ws, fmt, 27, 15, 33, f_k)?;
    // Zeile 29 (Row 28)
    write_rp_row(ws, fmt, 28, 16, 34, f_k)?;
    // Zeile 30 (Row 29)
    write_rp_row(ws, fmt, 29, 17, 35, f_k)?;
    // Zeile 31 (Row 30)
    write_rp_row(ws, fmt, 30, 18, 36, f_k)?;

    Ok(())
}

/// Schreibt eine Zeile des Right Panels
fn write_rp_row(
    ws: &mut Worksheet,
    fmt: &FormatMatrix,
    row: u32,
    left_num: u32,
    right_num: u32,
    f_k: &str,
) -> Result<(), XlsxError> {
    // === LINKE SEITE ===
    // J: Index
    if let Some(format) = fmt.get(row, 9) {
        ws.write_string_with_format(row, 9, &format!("{}. ", left_num), format)?;
    }
    // K: Text (Formel)
    if let Some(format) = fmt.get_locked(row, 10) {
        ws.write_formula_with_format(row, 10, f_k, &format)?;
    }
    // L: Date (blank)
    if let Some(format) = fmt.get(row, 11) {
        ws.write_blank(row, 11, format)?;
    }
    // M: Number (blank)
    if let Some(format) = fmt.get(row, 12) {
        ws.write_blank(row, 12, format)?;
    }
    // N: Number (blank)
    if let Some(format) = fmt.get(row, 13) {
        ws.write_blank(row, 13, format)?;
    }
    // O: Calc (Formel)
    let f_o = format!("=IF(M{}=\"\",\"\",N{}/M{})", row + 1, row + 1, row + 1);
    if let Some(format) = fmt.get_locked(row, 14) {
        ws.write_formula_with_format(row, 14, f_o.as_str(), &format)?;
    }

    // === RECHTE SEITE ===
    // Q: Index
    if let Some(format) = fmt.get(row, 16) {
        ws.write_string_with_format(row, 16, &format!("{}. ", right_num), format)?;
    }
    // R: Text (Formel)
    if let Some(format) = fmt.get_locked(row, 17) {
        ws.write_formula_with_format(row, 17, f_k, &format)?;
    }
    // S: Date (blank)
    if let Some(format) = fmt.get(row, 18) {
        ws.write_blank(row, 18, format)?;
    }
    // T: Number (blank)
    if let Some(format) = fmt.get(row, 19) {
        ws.write_blank(row, 19, format)?;
    }
    // U: Number (blank)
    if let Some(format) = fmt.get(row, 20) {
        ws.write_blank(row, 20, format)?;
    }
    // V: Calc (Formel)
    let f_v = format!("=IF(T{}=\"\",\"\",U{}/T{})", row + 1, row + 1, row + 1);
    if let Some(format) = fmt.get_locked(row, 21) {
        ws.write_formula_with_format(row, 21, f_v.as_str(), &format)?;
    }

    Ok(())
}
