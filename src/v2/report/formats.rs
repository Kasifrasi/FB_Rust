//! Zentrale Format-Definitionen für den Finanzbericht
//!
//! Enthält:
//! - ReportStyles: Basis-Styles (Farben, Borders, Standard-Formate)
//! - SectionStyles: Abgeleitete Styles für spezifische Bereiche
//! - FormatMatrix: Zell-spezifische Format-Zuordnung

use super::body::BodyLayout;
use rust_xlsxwriter::{Color, Format, FormatAlign, FormatBorder, FormatPattern};
use std::collections::HashMap;

// ============================================================================
// FormatMatrix: Zentrale Speicherung aller Zellformate
// ============================================================================

/// Zentrale Matrix für alle Zellformate
/// Default-Format ist unlocked (Arial 10), nur Formel-Zellen werden locked
pub struct FormatMatrix {
    formats: HashMap<(u32, u16), Format>,
}

impl FormatMatrix {
    pub fn new() -> Self {
        Self {
            formats: HashMap::new(),
        }
    }

    /// Setzt das Format für eine Zelle
    pub fn set(&mut self, row: u32, col: u16, format: &Format) {
        self.formats.insert((row, col), format.clone());
    }

    /// Holt das Format für eine Zelle (Standard - unlocked durch Default-Format)
    pub fn get(&self, row: u32, col: u16) -> Option<&Format> {
        self.formats.get(&(row, col))
    }

    /// Holt das Format für eine Zelle mit locked flag (für Formeln)
    pub fn get_locked(&self, row: u32, col: u16) -> Option<Format> {
        self.formats
            .get(&(row, col))
            .map(|f| f.clone().set_locked())
    }
}

impl Default for FormatMatrix {
    fn default() -> Self {
        Self::new()
    }
}

// ============================================================================
// ReportStyles: Basis-Styles
// ============================================================================

pub struct ReportStyles {
    // --- Colors ---
    pub fill_input: Color,
    pub fill_value: Color,
    pub fill_summary: Color,
    pub fl_orange: Color,
    pub blue_link: Color,

    // --- Borders ---
    pub border_thin: FormatBorder,
    pub border_medium: FormatBorder,
    pub border_dotted: FormatBorder,
    pub border_dashed: FormatBorder,

    // --- Base ---
    pub base: Format,

    // --- General Alignments ---
    pub left_center: Format,
    pub left_center_bold: Format,
    pub center_center: Format,
    pub center_center_bold: Format,
    pub center_center_bold_right: Format,
    pub left_top_wrap: Format,

    // --- Header Specifics ---
    pub header_label: Format,
    pub header_suffix: Format,
    pub input_orange_dashed: Format,
    pub header_right_box_top: Format,
    pub header_right_box_body: Format,
    pub link_style: Format,

    // --- Table Header ---
    pub table_header_base: Format,

    // --- Body ---
    pub body_label: Format,
    pub body_value: Format,
    pub body_input: Format,
    pub body_calc: Format,
    pub body_pct: Format,
    pub body_right: Format,

    // --- Summary Row ---
    pub summary_label: Format,
    pub summary_value: Format,
    pub summary_pct: Format,
    pub summary_right: Format,

    // --- Right Panel Base ---
    pub rp_index: Format,
    pub rp_text: Format,
    pub rp_date: Format,
    pub rp_number: Format,
    pub rp_calc: Format,
}

impl ReportStyles {
    pub fn new() -> Self {
        // Colors
        let fill_input = Color::RGB(0xFFFAE5);
        let fill_value = Color::RGB(0xF2F2F2);
        let fill_summary = Color::RGB(0xD9D9D9);
        let fl_orange = Color::RGB(0xFFCC00);
        let blue_link = Color::RGB(0x0563C1);

        // Borders
        let border_thin = FormatBorder::Thin;
        let border_medium = FormatBorder::Medium;
        let border_dotted = FormatBorder::Dotted;
        let border_dashed = FormatBorder::Dashed;

        // Base Format (Standard: Arial 10, unlocked)
        let base = Format::new()
            .set_font_name("Arial")
            .set_font_size(10.0)
            .set_unlocked();

        // Base Alignments
        let left_center = base
            .clone()
            .set_align(FormatAlign::Left)
            .set_align(FormatAlign::VerticalCenter);

        let left_center_bold = left_center.clone().set_bold();

        let center_center = base
            .clone()
            .set_align(FormatAlign::Center)
            .set_align(FormatAlign::VerticalCenter);

        let center_center_bold = center_center.clone().set_bold();
        let center_center_bold_right = center_center_bold.clone().set_border_right(border_thin);

        let left_top_wrap = base
            .clone()
            .set_align(FormatAlign::Left)
            .set_align(FormatAlign::Top)
            .set_text_wrap();

        // Header Specifics
        let header_label = left_center_bold.clone();
        let header_suffix = left_center.clone().set_font_color(Color::White);

        let input_orange_dashed = center_center
            .clone()
            .set_background_color(fl_orange)
            .set_border_bottom(border_dashed);

        let header_right_box_top = base
            .clone()
            .set_bold()
            .set_align(FormatAlign::Left)
            .set_border_top(border_medium)
            .set_border_left(border_medium);

        let header_right_box_body = left_top_wrap
            .clone()
            .set_border_left(border_medium)
            .set_border_right(border_medium);

        let link_style = base
            .clone()
            .set_align(FormatAlign::Left)
            .set_hyperlink()
            .set_border_left(border_medium)
            .set_border_right(border_medium)
            .set_border_bottom(border_medium);

        let table_header_base = base
            .clone()
            .set_align(FormatAlign::Center)
            .set_align(FormatAlign::VerticalCenter)
            .set_text_wrap()
            .set_border_top(border_medium);

        // Body Styles
        let body_label = base
            .clone()
            .set_bold()
            .set_align(FormatAlign::Left)
            .set_border_left(border_medium)
            .set_border_right(border_thin)
            .set_border_top(border_thin)
            .set_border_bottom(border_thin);

        let body_value = base
            .clone()
            .set_border(border_thin)
            .set_background_color(fill_value)
            .set_num_format("#,##0.00");

        let body_input = base
            .clone()
            .set_border(border_thin)
            .set_background_color(fill_input)
            .set_num_format("#,##0.00");

        let body_calc = base
            .clone()
            .set_border(border_thin)
            .set_background_color(fill_input)
            .set_num_format("#,##0.00");

        let body_pct = base.clone().set_border(border_thin).set_num_format("0%");

        let body_right = base
            .clone()
            .set_border(border_thin)
            .set_background_color(fill_input)
            .set_border_right(border_medium);

        // Summary Styles
        let summary_label = base
            .clone()
            .set_background_color(fill_summary)
            .set_bold()
            .set_align(FormatAlign::Left)
            .set_border_left(border_medium)
            .set_border_right(border_thin)
            .set_border_top(border_thin)
            .set_border_bottom(border_medium);

        let summary_value = base
            .clone()
            .set_background_color(fill_summary)
            .set_border(border_thin)
            .set_border_bottom(border_medium)
            .set_num_format("#,##0.00");

        let summary_pct = base
            .clone()
            .set_background_color(fill_summary)
            .set_border(border_thin)
            .set_border_bottom(border_medium)
            .set_num_format("0%");

        let summary_right = base
            .clone()
            .set_background_color(fill_summary)
            .set_border(border_thin)
            .set_border_bottom(border_medium)
            .set_border_right(border_medium)
            .set_num_format("#,##0.00");

        // Right Panel Base Styles
        let rp_index = base.clone().set_align(FormatAlign::Right);
        let rp_text = base.clone().set_align(FormatAlign::Left);

        let rp_date = base
            .clone()
            .set_align(FormatAlign::Center)
            .set_num_format("mm-dd-yy")
            .set_background_color(fill_input);

        let rp_number = base
            .clone()
            .set_align(FormatAlign::Right)
            .set_num_format("#,##0.00")
            .set_background_color(fill_input);

        let rp_calc = base
            .clone()
            .set_align(FormatAlign::Right)
            .set_num_format("0.0000");

        Self {
            fill_input,
            fill_value,
            fill_summary,
            fl_orange,
            blue_link,
            border_thin,
            border_medium,
            border_dotted,
            border_dashed,
            base,
            left_center,
            left_center_bold,
            center_center,
            center_center_bold,
            center_center_bold_right,
            left_top_wrap,
            header_label,
            header_suffix,
            input_orange_dashed,
            header_right_box_top,
            header_right_box_body,
            link_style,
            table_header_base,
            body_label,
            body_value,
            body_input,
            body_calc,
            body_pct,
            body_right,
            summary_label,
            summary_value,
            summary_pct,
            summary_right,
            rp_index,
            rp_text,
            rp_date,
            rp_number,
            rp_calc,
        }
    }
}

impl Default for ReportStyles {
    fn default() -> Self {
        Self::new()
    }
}

// ============================================================================
// SectionStyles: Abgeleitete Styles für spezifische Bereiche
// ============================================================================

pub struct SectionStyles {
    // Header Section (Row 0-9)
    pub fmt_top_med: Format,
    pub fmt_top_right_med: Format,
    pub fmt_d5: Format,
    pub fmt_d6: Format,
    pub fmt_row7_base: Format,
    pub fmt_row7_date: Format,
    pub orange_dotted: Format,
    pub value_dotted: Format,
    pub input_dotted: Format,

    // Table Header Section (Row 10-13)
    pub th_b: Format,
    pub th_c_top: Format,
    pub th_d: Format,
    pub th_d_bold: Format,
    pub th_h: Format,
    pub th_side: Format,
    pub th_side_bold: Format,
    pub th_bot_side: Format,
    pub th_bot_right: Format,

    // Table Body Section (Row 14-19)
    pub body_label_top: Format,
    pub body_label_no_h: Format,

    // Right Panel Header (Row 10-12)
    pub rp_hdr_left: Format,
    pub rp_hdr_mid: Format,
    pub rp_hdr_mid_top: Format,
    pub rp_hdr_right: Format,
    pub rp_hdr_right_top: Format,

    // Right Panel Body (Row 13-30)
    pub rp_idx: Format,
    pub rp_idx_last: Format,
    pub rp_txt: Format,
    pub rp_txt_last: Format,
    pub rp_date: Format,
    pub rp_date_last: Format,
    pub rp_num: Format,
    pub rp_num_last: Format,
    pub rp_calc: Format,
    pub rp_calc_last: Format,
}

impl SectionStyles {
    pub fn new(s: &ReportStyles) -> Self {
        // Header Section
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

        // Table Header Section
        let th_b = s.table_header_base.clone().set_border_left(s.border_medium);
        let th_c_top = s.table_header_base.clone().set_border_right(s.border_thin);

        let th_d = s
            .table_header_base
            .clone()
            .set_border_left(s.border_thin)
            .set_border_right(s.border_thin);

        let th_d_bold = th_d.clone().set_bold();

        let th_h = s
            .table_header_base
            .clone()
            .set_border_left(s.border_thin)
            .set_border_right(s.border_medium);

        let th_side = s
            .base
            .clone()
            .set_align(FormatAlign::Center)
            .set_align(FormatAlign::VerticalCenter)
            .set_border_left(s.border_medium)
            .set_border_right(s.border_thin);

        let th_side_bold = th_side.clone().set_bold();

        let th_bot_side = s
            .base
            .clone()
            .set_align(FormatAlign::Center)
            .set_align(FormatAlign::VerticalCenter)
            .set_border_left(s.border_medium)
            .set_border_bottom(s.border_thin);

        let th_bot_right = s
            .base
            .clone()
            .set_border_bottom(s.border_thin)
            .set_border_right(s.border_thin);

        // Table Body Section
        let body_label_top = s.body_label.clone().set_border_bottom(FormatBorder::None);
        let body_label_no_h = s
            .body_label
            .clone()
            .set_border_top(FormatBorder::None)
            .set_border_bottom(FormatBorder::None);

        // Right Panel Header
        let rp_hdr_left = s.base.clone().set_border_left(s.border_thin);
        let rp_hdr_mid = s.base.clone();
        let rp_hdr_mid_top = rp_hdr_mid.clone().set_border_top(s.border_thin);
        let rp_hdr_right = s.base.clone().set_border_right(s.border_thin);
        let rp_hdr_right_top = rp_hdr_right.clone().set_border_top(s.border_thin);

        // Right Panel Body
        let rp_idx = s.rp_index.clone().set_border_left(s.border_thin);
        let rp_idx_last = rp_idx.clone().set_border_bottom(s.border_thin);
        let rp_txt = s.rp_text.clone().set_border_right(s.border_thin);
        let rp_txt_last = rp_txt.clone().set_border_bottom(s.border_thin);

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
            orange_dotted,
            value_dotted,
            input_dotted,
            th_b,
            th_c_top,
            th_d,
            th_d_bold,
            th_h,
            th_side,
            th_side_bold,
            th_bot_side,
            th_bot_right,
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
// build_format_matrix: Zentrale Zell-Format-Zuordnung
// ============================================================================

pub fn build_format_matrix(styles: &ReportStyles, sec: &SectionStyles) -> FormatMatrix {
    let mut m = FormatMatrix::new();

    // ========================================================================
    // HEADER BEREICH (Rows 0-9)
    // ========================================================================

    // Row 0
    m.set(0, 1, &styles.header_label);
    m.set(0, 9, &styles.header_right_box_top);
    m.set(0, 10, &sec.fmt_top_med);
    m.set(0, 11, &sec.fmt_top_med);
    m.set(0, 12, &sec.fmt_top_med);
    m.set(0, 13, &sec.fmt_top_med);
    m.set(0, 14, &sec.fmt_top_right_med);

    // Row 1
    m.set(1, 1, &styles.header_suffix);
    m.set(1, 3, &styles.left_center);
    m.set(1, 4, &styles.input_orange_dashed);
    m.set(1, 9, &styles.header_right_box_body);

    // Row 2
    m.set(2, 1, &styles.header_label);
    m.set(2, 3, &styles.left_center);
    m.set(2, 4, &styles.input_orange_dashed);

    // Row 3
    m.set(3, 9, &styles.link_style);

    // Row 4
    m.set(4, 1, &styles.left_center);
    m.set(4, 3, &sec.fmt_d5);

    // Row 5
    m.set(5, 1, &styles.left_center);
    m.set(5, 3, &sec.fmt_d6);

    // Row 6
    m.set(6, 9, &sec.orange_dotted);

    // Row 7
    m.set(7, 1, &styles.left_center);
    m.set(7, 3, &sec.fmt_row7_base);
    m.set(7, 4, &sec.fmt_row7_date);
    m.set(7, 5, &sec.fmt_row7_base);
    m.set(7, 6, &sec.fmt_row7_date);
    m.set(7, 9, &sec.value_dotted);

    // Row 8
    m.set(8, 1, &styles.left_center);
    m.set(8, 3, &sec.fmt_row7_base);
    m.set(8, 4, &sec.fmt_row7_date);
    m.set(8, 5, &sec.fmt_row7_base);
    m.set(8, 6, &sec.fmt_row7_date);
    m.set(8, 9, &sec.input_dotted);

    // ========================================================================
    // TABLE HEADER (Rows 10-13)
    // ========================================================================

    // Row 10
    m.set(10, 1, &sec.th_b);
    m.set(10, 2, &sec.th_c_top);
    m.set(10, 3, &sec.th_d);
    m.set(10, 4, &sec.th_d_bold);
    m.set(10, 5, &sec.th_d);
    m.set(10, 6, &sec.th_d);
    m.set(10, 7, &sec.th_h);

    // Row 11
    m.set(11, 1, &sec.th_side_bold);

    // Row 12
    m.set(12, 1, &sec.th_side);
    m.set(12, 11, &styles.center_center_bold);
    m.set(12, 12, &styles.center_center_bold);
    m.set(12, 13, &styles.center_center_bold);
    m.set(12, 14, &styles.center_center_bold_right);
    m.set(12, 18, &styles.center_center_bold);
    m.set(12, 19, &styles.center_center_bold);
    m.set(12, 20, &styles.center_center_bold);
    m.set(12, 21, &styles.center_center_bold_right);

    // Row 13
    m.set(13, 1, &sec.th_bot_side);
    m.set(13, 2, &sec.th_bot_right);

    // ========================================================================
    // TABLE BODY (Rows 14-18)
    // ========================================================================

    // Row 14 (erste Datenzeile)
    m.set(14, 1, &sec.body_label_top);
    m.set(14, 3, &styles.body_value);
    m.set(14, 4, &styles.body_input);
    m.set(14, 5, &styles.body_calc);
    m.set(14, 6, &styles.body_pct);
    m.set(14, 7, &styles.body_right);

    // Rows 15-18
    for row in 15..=18 {
        m.set(row, 1, &sec.body_label_no_h);
        m.set(row, 3, &styles.body_value);
        m.set(row, 4, &styles.body_input);
        m.set(row, 5, &styles.body_calc);
        m.set(row, 6, &styles.body_pct);
        m.set(row, 7, &styles.body_right);
    }

    // ========================================================================
    // SUMMARY ROW (Row 19)
    // ========================================================================

    m.set(19, 1, &styles.summary_label);
    m.set(19, 3, &styles.summary_value);
    m.set(19, 4, &styles.summary_value);
    m.set(19, 5, &styles.summary_value);
    m.set(19, 6, &styles.summary_pct);
    m.set(19, 7, &styles.summary_right);

    // ========================================================================
    // RIGHT PANEL (Row 10-30, Cols J-O und Q-V)
    // ========================================================================

    // Row 10 (J11-O11, Q11-V11): Header mit top border
    let j11_fmt = styles
        .left_center_bold
        .clone()
        .set_border_left(styles.border_thin)
        .set_border_top(styles.border_thin);
    let q11_fmt = j11_fmt.clone();

    m.set(10, 9, &j11_fmt);
    m.set(10, 11, &sec.rp_hdr_mid_top);
    m.set(10, 12, &sec.rp_hdr_mid_top);
    m.set(10, 13, &sec.rp_hdr_mid_top);
    m.set(10, 14, &sec.rp_hdr_right_top);
    m.set(10, 16, &q11_fmt);
    m.set(10, 18, &sec.rp_hdr_mid_top);
    m.set(10, 19, &sec.rp_hdr_mid_top);
    m.set(10, 20, &sec.rp_hdr_mid_top);
    m.set(10, 21, &sec.rp_hdr_right_top);

    // Row 11 (J12-O12, Q12-V12): Header ohne top border
    m.set(11, 9, &sec.rp_hdr_left);
    m.set(11, 10, &sec.rp_hdr_mid);
    m.set(11, 11, &sec.rp_hdr_mid);
    m.set(11, 12, &sec.rp_hdr_mid);
    m.set(11, 13, &sec.rp_hdr_mid);
    m.set(11, 14, &sec.rp_hdr_right);
    m.set(11, 16, &sec.rp_hdr_left);
    m.set(11, 17, &sec.rp_hdr_mid);
    m.set(11, 18, &sec.rp_hdr_mid);
    m.set(11, 19, &sec.rp_hdr_mid);
    m.set(11, 20, &sec.rp_hdr_mid);
    m.set(11, 21, &sec.rp_hdr_right);

    // Row 12 (J13, K13, Q13, R13): nur blanks, L13-O13 und S13-V13 haben center_center_bold
    m.set(12, 9, &sec.rp_hdr_left);
    m.set(12, 10, &sec.rp_hdr_mid);
    m.set(12, 16, &sec.rp_hdr_left);
    m.set(12, 17, &sec.rp_hdr_mid);

    // Rows 13-29: Body ohne bottom border
    for row in 13..=29 {
        m.set(row, 9, &sec.rp_idx);
        m.set(row, 10, &sec.rp_txt);
        m.set(row, 11, &sec.rp_date);
        m.set(row, 12, &sec.rp_num);
        m.set(row, 13, &sec.rp_num);
        m.set(row, 14, &sec.rp_calc);
        m.set(row, 16, &sec.rp_idx);
        m.set(row, 17, &sec.rp_txt);
        m.set(row, 18, &sec.rp_date);
        m.set(row, 19, &sec.rp_num);
        m.set(row, 20, &sec.rp_num);
        m.set(row, 21, &sec.rp_calc);
    }

    // Row 30: Body mit bottom border
    m.set(30, 9, &sec.rp_idx_last);
    m.set(30, 10, &sec.rp_txt_last);
    m.set(30, 11, &sec.rp_date_last);
    m.set(30, 12, &sec.rp_num_last);
    m.set(30, 13, &sec.rp_num_last);
    m.set(30, 14, &sec.rp_calc_last);
    m.set(30, 16, &sec.rp_idx_last);
    m.set(30, 17, &sec.rp_txt_last);
    m.set(30, 18, &sec.rp_date_last);
    m.set(30, 19, &sec.rp_num_last);
    m.set(30, 20, &sec.rp_num_last);
    m.set(30, 21, &sec.rp_calc_last);

    m
}

// ============================================================================
// BodyStyles: Formate für den dynamischen Body-Bereich
// ============================================================================

/// Styles für den dynamischen Body-Bereich (Kostenkategorien)
pub struct BodyStyles {
    // === Category Header (bold, keine Füllung) ===
    /// B: Kategorie-Nummer ("1."), rechts, bold, medium left border
    pub cat_header_b: Format,
    /// C: VLOOKUP Label, links, bold
    pub cat_header_c: Format,
    /// D-F: Werte-Spalten (leer im Header)
    pub cat_header_value: Format,
    /// G: Prozent-Spalte (leer im Header)
    pub cat_header_pct: Format,
    /// H: Bemerkung, medium right border
    pub cat_header_h: Format,

    // === Position Row (Input Cells, gelb) ===
    /// B: Positions-Nummer ("1.1"), rechts
    pub pos_b: Format,
    /// C: Beschreibung, links, input fill, wrap
    pub pos_c: Format,
    /// D: Bewilligt, grau fill (berechnet oder value)
    pub pos_d: Format,
    /// E-F: Einnahmen, input fill
    pub pos_ef: Format,
    /// G: Prozent (Formel)
    pub pos_g: Format,
    /// H: Bemerkung, input fill, wrap, medium right
    pub pos_h: Format,

    // === Category Footer (grau fill, bold) ===
    /// B: Sum-Label via VLOOKUP
    pub footer_b: Format,
    /// C: leer
    pub footer_c: Format,
    /// D-F: SUMPRODUCT Formeln
    pub footer_value: Format,
    /// G: Ratio-Formel
    pub footer_pct: Format,
    /// H: leer, medium right
    pub footer_h: Format,

    // === Single-Row Category (wie Position, aber mit bold Label) ===
    /// B: Kategorie-Nummer ("6."), rechts, bold
    pub single_b: Format,
    /// C: VLOOKUP Label, bold
    pub single_c: Format,
    // D-H: gleich wie pos_*

    // === Total Row (grau fill, bold, medium bottom) ===
    /// B: "Gesamt" Label
    pub total_b: Format,
    /// C: leer
    pub total_c: Format,
    /// D-F: SUM Formeln
    pub total_value: Format,
    /// G: Ratio-Formel
    pub total_pct: Format,
    /// H: leer, medium right + bottom
    pub total_h: Format,
}

impl BodyStyles {
    pub fn new(s: &ReportStyles) -> Self {
        let thin = s.border_thin;
        let medium = s.border_medium;

        // === Category Header ===
        let cat_header_b = s
            .base
            .clone()
            .set_align(FormatAlign::Right)
            .set_bold()
            .set_num_format("@")
            .set_border_left(medium)
            .set_border_top(thin)
            .set_border_bottom(thin)
            .set_border_right(thin);

        let cat_header_c = s
            .base
            .clone()
            .set_align(FormatAlign::Left)
            .set_bold()
            .set_border(thin);

        let cat_header_value = s
            .base
            .clone()
            .set_align(FormatAlign::Right)
            .set_bold()
            .set_text_wrap()
            .set_num_format("#,##0.00")
            .set_border(thin);

        let cat_header_pct = s
            .base
            .clone()
            .set_align(FormatAlign::Right)
            .set_bold()
            .set_num_format("0%")
            .set_border(thin);

        let cat_header_h = s
            .base
            .clone()
            .set_align(FormatAlign::Left)
            .set_bold()
            .set_border_left(thin)
            .set_border_top(thin)
            .set_border_bottom(thin)
            .set_border_right(medium);

        // === Position Row ===
        let pos_b = s
            .base
            .clone()
            .set_align(FormatAlign::Right)
            .set_num_format("@")
            .set_border_left(medium)
            .set_border_top(thin)
            .set_border_bottom(thin)
            .set_border_right(thin);

        let pos_c = s
            .base
            .clone()
            .set_background_color(s.fill_input)
            .set_text_wrap()
            .set_border(thin);

        let pos_d = s
            .base
            .clone()
            .set_background_color(s.fill_value)
            .set_text_wrap()
            .set_num_format("#,##0.00")
            .set_border(thin);

        let pos_ef = s
            .base
            .clone()
            .set_background_color(s.fill_input)
            .set_text_wrap()
            .set_num_format("#,##0.00")
            .set_border(thin);

        let pos_g = s
            .base
            .clone()
            .set_align(FormatAlign::Right)
            .set_num_format("0%")
            .set_border(thin);

        let pos_h = s
            .base
            .clone()
            .set_background_color(s.fill_input)
            .set_text_wrap()
            .set_border_left(thin)
            .set_border_top(thin)
            .set_border_bottom(thin)
            .set_border_right(medium);

        // === Category Footer (grau) ===
        let gray_bold = s
            .base
            .clone()
            .set_foreground_color(s.fill_summary)
            .set_pattern(FormatPattern::Solid)
            .set_bold();

        let footer_b = gray_bold
            .clone()
            .set_border_left(medium)
            .set_border_top(thin)
            .set_border_bottom(thin)
            .set_border_right(thin);

        let footer_c = gray_bold.clone().set_border(thin);

        let footer_value = gray_bold
            .clone()
            .set_align(FormatAlign::Right)
            .set_num_format("#,##0.00")
            .set_border(thin);

        let footer_pct = gray_bold
            .clone()
            .set_align(FormatAlign::Right)
            .set_num_format("0%")
            .set_border(thin);

        let footer_h = gray_bold
            .clone()
            .set_border_left(thin)
            .set_border_top(thin)
            .set_border_bottom(thin)
            .set_border_right(medium);

        // === Single-Row Category ===
        let single_b = s
            .base
            .clone()
            .set_align(FormatAlign::Right)
            .set_bold()
            .set_num_format("@")
            .set_border_left(medium)
            .set_border_top(thin)
            .set_border_bottom(thin)
            .set_border_right(thin);

        // Für Kategorie 8: grauer Text
        let single_c = s
            .base
            .clone()
            .set_align(FormatAlign::Left)
            .set_bold()
            .set_border(thin);

        // === Total Row (grau, medium bottom) ===
        let total_b = gray_bold
            .clone()
            .set_border_left(medium)
            .set_border_top(thin)
            .set_border_bottom(medium)
            .set_border_right(thin);

        let total_c = gray_bold
            .clone()
            .set_border_top(thin)
            .set_border_bottom(medium)
            .set_border_left(thin)
            .set_border_right(thin);

        let total_value = gray_bold
            .clone()
            .set_align(FormatAlign::Right)
            .set_num_format("#,##0.00")
            .set_border_top(thin)
            .set_border_bottom(medium)
            .set_border_left(thin)
            .set_border_right(thin);

        let total_pct = gray_bold
            .clone()
            .set_align(FormatAlign::Right)
            .set_num_format("0%")
            .set_border_top(thin)
            .set_border_bottom(medium)
            .set_border_left(thin)
            .set_border_right(thin);

        let total_h = gray_bold
            .clone()
            .set_border_left(thin)
            .set_border_top(thin)
            .set_border_bottom(medium)
            .set_border_right(medium);

        Self {
            cat_header_b,
            cat_header_c,
            cat_header_value,
            cat_header_pct,
            cat_header_h,
            pos_b,
            pos_c,
            pos_d,
            pos_ef,
            pos_g,
            pos_h,
            footer_b,
            footer_c,
            footer_value,
            footer_pct,
            footer_h,
            single_b,
            single_c,
            total_b,
            total_c,
            total_value,
            total_pct,
            total_h,
        }
    }
}

// ============================================================================
// extend_format_matrix_with_body: Fügt Body-Formate zur Matrix hinzu
// ============================================================================

/// Erweitert die FormatMatrix um die dynamischen Body-Formate
pub fn extend_format_matrix_with_body(
    m: &mut FormatMatrix,
    styles: &ReportStyles,
    layout: &BodyLayout,
) {
    let body = BodyStyles::new(styles);

    for cat in &layout.categories {
        // === Multi-Row Kategorie ===
        if let Some(header_row) = cat.header_row {
            // Header-Zeile
            m.set(header_row, 1, &body.cat_header_b); // B
            m.set(header_row, 2, &body.cat_header_c); // C
            m.set(header_row, 3, &body.cat_header_value); // D
            m.set(header_row, 4, &body.cat_header_value); // E
            m.set(header_row, 5, &body.cat_header_value); // F
            m.set(header_row, 6, &body.cat_header_pct); // G
            m.set(header_row, 7, &body.cat_header_h); // H
        }

        if let Some(positions) = &cat.positions {
            // Positions-Zeilen
            for row in positions.start_row..=positions.end_row {
                m.set(row, 1, &body.pos_b); // B
                m.set(row, 2, &body.pos_c); // C
                m.set(row, 3, &body.pos_d); // D
                m.set(row, 4, &body.pos_ef); // E
                m.set(row, 5, &body.pos_ef); // F
                m.set(row, 6, &body.pos_g); // G
                m.set(row, 7, &body.pos_h); // H
            }
        }

        if let Some(footer_row) = cat.footer_row {
            // Footer-Zeile
            m.set(footer_row, 1, &body.footer_b); // B
            m.set(footer_row, 2, &body.footer_c); // C
            m.set(footer_row, 3, &body.footer_value); // D
            m.set(footer_row, 4, &body.footer_value); // E
            m.set(footer_row, 5, &body.footer_value); // F
            m.set(footer_row, 6, &body.footer_pct); // G
            m.set(footer_row, 7, &body.footer_h); // H
        }

        // === Single-Row Kategorie ===
        if let Some(single_row) = cat.single_row {
            // Single-Row verwendet Position-Formate für D-H, aber bold B und C
            m.set(single_row, 1, &body.single_b); // B

            // Kategorie 8 hat grauen Text
            if cat.meta.num == 8 {
                let gray_text = body.single_c.clone().set_font_color(Color::RGB(0xBFBFBF));
                m.set(single_row, 2, &gray_text);
            } else {
                m.set(single_row, 2, &body.single_c);
            }

            m.set(single_row, 3, &body.pos_d); // D
            m.set(single_row, 4, &body.pos_ef); // E
            m.set(single_row, 5, &body.pos_ef); // F
            m.set(single_row, 6, &body.pos_g); // G
            m.set(single_row, 7, &body.pos_h); // H
        }
    }

    // === Total-Zeile ===
    m.set(layout.total_row, 1, &body.total_b); // B
    m.set(layout.total_row, 2, &body.total_c); // C
    m.set(layout.total_row, 3, &body.total_value); // D
    m.set(layout.total_row, 4, &body.total_value); // E
    m.set(layout.total_row, 5, &body.total_value); // F
    m.set(layout.total_row, 6, &body.total_pct); // G
    m.set(layout.total_row, 7, &body.total_h); // H
}
