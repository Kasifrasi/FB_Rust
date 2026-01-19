use rust_xlsxwriter::{Color, Format, FormatAlign, FormatPattern, FormatUnderline};

/// Central style palette translating OpenPyXL helper constants to rust_xlsxwriter `Format`s.
#[derive(Debug, Clone)]
pub struct StylePalette {
    pub font_normal: Format,
    pub font_bold: Format,
    pub font_gray_bold: Format,
    pub font_gray: Format,
    pub align_left: Format,
    pub align_left_wrap: Format,
    pub align_center: Format,
    pub align_center_wrap: Format,
    pub align_right: Format,
    pub align_right_wrap: Format,
    pub align_percent: Format,
    pub fill_input: Format,
    pub fill_value: Format,
    pub fill_summary: Format,
    pub fill_white: Format,
}

impl Default for StylePalette {
    fn default() -> Self {
        Self::new()
    }
}

impl StylePalette {
    /// Create a new palette with all predefined formats.
    pub fn new() -> Self {
        let font_normal = Format::new().set_font_name("Arial").set_font_size(10.0);

        let font_bold = font_normal.clone().set_bold();

        let font_gray = font_normal.clone().set_font_color(Color::RGB(0xBFBFBF));

        let font_gray_bold = font_gray.clone().set_bold();

        let align_left = font_normal
            .clone()
            .set_align(FormatAlign::Left)
            .set_align(FormatAlign::VerticalCenter);
        let align_left_wrap = align_left.clone().set_text_wrap();

        let align_center = font_normal
            .clone()
            .set_align(FormatAlign::Center)
            .set_align(FormatAlign::VerticalCenter);
        let align_center_wrap = align_center.clone().set_text_wrap();

        let align_right = font_normal
            .clone()
            .set_align(FormatAlign::Right)
            .set_align(FormatAlign::VerticalCenter);
        let align_right_wrap = align_right.clone().set_text_wrap();

        let align_percent = align_right.clone();

        let fill_input = align_left
            .clone()
            .set_foreground_color(Color::RGB(0xFFFAE5))
            .set_pattern(FormatPattern::Solid);

        let fill_value = align_right
            .clone()
            .set_foreground_color(Color::RGB(0xF2F2F2))
            .set_pattern(FormatPattern::Solid);

        let fill_summary = align_left
            .clone()
            .set_foreground_color(Color::RGB(0xD9D9D9))
            .set_pattern(FormatPattern::Solid);

        let fill_white = align_left
            .clone()
            .set_foreground_color(Color::RGB(0xFFFFFF))
            .set_pattern(FormatPattern::Solid);

        Self {
            font_normal,
            font_bold,
            font_gray_bold,
            font_gray,
            align_left,
            align_left_wrap,
            align_center,
            align_center_wrap,
            align_right,
            align_right_wrap,
            align_percent,
            fill_input,
            fill_value,
            fill_summary,
            fill_white,
        }
    }

    /// Returns a clone of the given base format with underline applied.
    pub fn underline(&self, base: &Format, underline: FormatUnderline) -> Format {
        base.clone().set_underline(underline)
    }
}
