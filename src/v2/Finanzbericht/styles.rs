use rust_xlsxwriter::{Color, Format, FormatAlign, FormatBorder};

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

    // --- General Alignments ---
    pub left_center: Format,
    pub left_center_bold: Format,
    pub center_center: Format,
    pub center_center_bold: Format,
    pub left_top_wrap: Format,

    // --- Header Specifics ---
    pub header_label: Format,          // Left, VCenter, Bold
    pub header_suffix: Format,         // Left, VCenter, White Text
    pub input_orange_dashed: Format,   // Center, VCenter, Orange BG, Dashed Bottom
    pub header_right_box_top: Format,  // J1 Top/Left Medium
    pub header_right_box_body: Format, // J2:O3 Left/Right Medium, Top Align, Wrap
    pub link_style: Format,            // Blue, Left, Left/Right/Bottom Medium

    // --- Table Header ---
    // The table header (Row 10-13) has complex borders.
    // We define the base style (Center, VCenter, Wrap, Top Medium).
    pub table_header_base: Format,

    // --- Base ---
    // Exposed so local styles can derive from it
    pub base: Format,

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

        // Base Format (Standard: Arial 10)
        let base = Format::new().set_font_name("Arial").set_font_size(10.0);

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
            .set_font_color(blue_link)
            .set_border_left(border_medium)
            .set_border_right(border_medium)
            .set_border_bottom(border_medium);

        let table_header_base = base
            .clone()
            .set_align(FormatAlign::Center)
            .set_align(FormatAlign::VerticalCenter)
            .set_text_wrap()
            .set_border_top(border_medium);

        // --- Body Styles ---
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

        // --- Summary Styles ---
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
            left_center,
            left_center_bold,
            center_center,
            center_center_bold,
            left_top_wrap,
            header_label,
            header_suffix,
            input_orange_dashed,
            header_right_box_top,
            header_right_box_body,
            link_style,
            table_header_base,
            base,
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
        }
    }
}
