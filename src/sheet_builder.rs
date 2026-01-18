use regex::Regex;
use rust_xlsxwriter::*;
use std::collections::HashMap;

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
    if let Some(idx) = col_index(col) {
        // Ignore errors; hiding non-existent columns is non-fatal in this context.
        let _ = ws.set_column_hidden(idx);
    }
}

/// Map column letter (A-Z) to zero-based index.
fn col_index(col: char) -> Option<u16> {
    let upper = col.to_ascii_uppercase();
    if ('A'..='Z').contains(&upper) {
        Some((upper as u8 - b'A') as u16)
    } else {
        None
    }
}

// ---------------------------------------------------------------------------
// Virtual Sheet Model
// ---------------------------------------------------------------------------

#[derive(Clone, Debug, Default)]
struct CellStyle {
    font_name: Option<String>,
    font_size: Option<f64>,
    bold: Option<bool>,
    italic: Option<bool>,
    underline: Option<FormatUnderline>,
    font_color: Option<Color>,

    fill_color: Option<Color>,
    fill_pattern: Option<FormatPattern>,

    align_h: Option<FormatAlign>,
    align_v: Option<FormatAlign>,
    wrap_text: Option<bool>,

    border_left: Option<FormatBorder>,
    border_right: Option<FormatBorder>,
    border_top: Option<FormatBorder>,
    border_bottom: Option<FormatBorder>,

    num_format: Option<String>,
}

#[derive(Clone, Debug)]
pub enum CellValue {
    String(String),
    Number(f64),
    Formula(String),
    Empty,
}

struct VirtualCell {
    value: CellValue,
    style: CellStyle,
}

impl Default for VirtualCell {
    fn default() -> Self {
        Self {
            value: CellValue::Empty,
            style: CellStyle::default(),
        }
    }
}

pub struct SheetModel {
    cells: HashMap<(u32, u16), VirtualCell>,
    merges: Vec<(u32, u16, u32, u16)>, // (row1, col1, row2, col2)
    col_widths: HashMap<u16, f64>,
    row_heights: HashMap<u32, f64>,
    hidden_cols: Vec<u16>,
    grouped_cols: Vec<(u16, u16)>,
}

impl SheetModel {
    pub fn new() -> Self {
        Self {
            cells: HashMap::new(),
            merges: Vec::new(),
            col_widths: HashMap::new(),
            row_heights: HashMap::new(),
            hidden_cols: Vec::new(),
            grouped_cols: Vec::new(),
        }
    }

    fn get_cell_mut(&mut self, row: u32, col: u16) -> &mut VirtualCell {
        self.cells.entry((row, col)).or_default()
    }

    pub fn set_col_width(&mut self, col: u16, width: f64) {
        self.col_widths.insert(col, width);
    }

    pub fn set_row_height(&mut self, row: u32, height: f64) {
        self.row_heights.insert(row, height);
    }

    pub fn group_columns(&mut self, start_col: u16, end_col: u16) {
        self.grouped_cols.push((start_col, end_col));
    }

    pub fn _hide_column(&mut self, col: u16) {
        self.hidden_cols.push(col);
    }

    pub fn merge_cells(&mut self, range: &str) {
        if let Some((r1, c1, r2, c2)) = parse_range(range) {
            self.merges.push((r1, c1, r2, c2));
        }
    }

    // --- Formatting Helpers ---

    pub fn apply_alignment(
        &mut self,
        ranges: &[&str],
        align_h: Option<FormatAlign>,
        align_v: Option<FormatAlign>,
        wrap: Option<bool>,
    ) {
        for range in ranges {
            if let Some((r1, c1, r2, c2)) = parse_range(range) {
                for r in r1..=r2 {
                    for c in c1..=c2 {
                        let cell = self.get_cell_mut(r, c);
                        if let Some(h) = align_h {
                            cell.style.align_h = Some(h);
                        }
                        if let Some(v) = align_v {
                            cell.style.align_v = Some(v);
                        }
                        if let Some(w) = wrap {
                            cell.style.wrap_text = Some(w);
                        }
                    }
                }
            }
        }
    }

    pub fn apply_number_format(&mut self, ranges: &[&str], format: &str) {
        for range in ranges {
            if let Some((r1, c1, r2, c2)) = parse_range(range) {
                for r in r1..=r2 {
                    for c in c1..=c2 {
                        self.get_cell_mut(r, c).style.num_format = Some(format.to_string());
                    }
                }
            }
        }
    }

    pub fn apply_font(
        &mut self,
        ranges: &[&str],
        bold: Option<bool>,
        color: Option<Color>,
        name: Option<&str>,
        size: Option<f64>,
    ) {
        for range in ranges {
            if let Some((r1, c1, r2, c2)) = parse_range(range) {
                for r in r1..=r2 {
                    for c in c1..=c2 {
                        let cell = self.get_cell_mut(r, c);
                        if let Some(b) = bold {
                            cell.style.bold = Some(b);
                        }
                        if let Some(c) = color {
                            cell.style.font_color = Some(c);
                        }
                        if let Some(n) = name {
                            cell.style.font_name = Some(n.to_string());
                        }
                        if let Some(s) = size {
                            cell.style.font_size = Some(s);
                        }
                    }
                }
            }
        }
    }

    pub fn apply_fill(&mut self, ranges: &[&str], color: Color) {
        for range in ranges {
            if let Some((r1, c1, r2, c2)) = parse_range(range) {
                for r in r1..=r2 {
                    for c in c1..=c2 {
                        let cell = self.get_cell_mut(r, c);
                        cell.style.fill_color = Some(color);
                        cell.style.fill_pattern = Some(FormatPattern::Solid);
                    }
                }
            }
        }
    }

    pub fn apply_grid_border(
        &mut self,
        ranges: &[&str],
        outer_sides: &[&str],
        inner_lines: &str,
        style: FormatBorder,
        _color: Color,
    ) {
        // Note: Ignoring color for now (default black), as rust_xlsxwriter border color setting is verbose.
        // Assuming black border.

        let has_all = outer_sides.contains(&"all");
        let top_out = has_all || outer_sides.contains(&"top");
        let bottom_out = has_all || outer_sides.contains(&"bottom");
        let left_out = has_all || outer_sides.contains(&"left");
        let right_out = has_all || outer_sides.contains(&"right");

        // inner_lines: "both", "horizontal", "vertical", "none"
        let inner_h = inner_lines == "both" || inner_lines == "horizontal";
        let inner_v = inner_lines == "both" || inner_lines == "vertical";

        for range in ranges {
            if let Some((r1, c1, r2, c2)) = parse_range(range) {
                for r in r1..=r2 {
                    for c in c1..=c2 {
                        let cell = self.get_cell_mut(r, c);

                        let is_top = r == r1;
                        let is_bottom = r == r2;
                        let is_left = c == c1;
                        let is_right = c == c2;

                        // Top
                        if is_top {
                            if top_out {
                                cell.style.border_top = Some(style);
                            }
                        } else if inner_h {
                            cell.style.border_top = Some(style);
                        }

                        // Bottom
                        if is_bottom {
                            if bottom_out {
                                cell.style.border_bottom = Some(style);
                            }
                        } else if inner_h {
                            cell.style.border_bottom = Some(style);
                        }

                        // Left
                        if is_left {
                            if left_out {
                                cell.style.border_left = Some(style);
                            }
                        } else if inner_v {
                            cell.style.border_left = Some(style);
                        }

                        // Right
                        if is_right {
                            if right_out {
                                cell.style.border_right = Some(style);
                            }
                        } else if inner_v {
                            cell.style.border_right = Some(style);
                        }
                    }
                }
            }
        }
    }

    pub fn set_value(&mut self, cell_ref: &str, val: CellValue) {
        if let Some((r, c, _, _)) = parse_range(cell_ref) {
            let cell = self.get_cell_mut(r, c);
            cell.value = val;
        }
    }

    pub fn set_string(&mut self, cell_ref: &str, val: &str) {
        self.set_value(cell_ref, CellValue::String(val.to_string()));
    }

    pub fn set_number(&mut self, cell_ref: &str, val: f64) {
        self.set_value(cell_ref, CellValue::Number(val));
    }

    pub fn set_formula(&mut self, cell_ref: &str, val: &str) {
        self.set_value(cell_ref, CellValue::Formula(val.to_string()));
    }

    pub fn apply_ratio_formulas(&mut self, ranges: Vec<(&str, &str, &str, u32, &str)>) {
        for (range, col_num, col_den, max_row, error_val) in ranges {
            // range: "O14:O31"
            if let Some((r1, c1, r2, _c2)) = parse_range(range) {
                for r in r1..=r2 {
                    let error_str = if error_val.is_empty() {
                        "\"\""
                    } else {
                        error_val
                    };
                    let f = format!("=IFERROR(INDEX(${}$1:${}${max_row},ROW())/INDEX(${}$1:${}${max_row},ROW()),{})",
                         col_num, col_num, col_den, col_den, error_str);

                    // c1 is the column index.
                    for c in c1..=c1 {
                        // Usually single column
                        let cell = self.get_cell_mut(r, c);
                        cell.value = CellValue::Formula(f.clone());
                    }
                }
            }
        }
    }

    pub fn apply_vlookup_formulas(&mut self, items: Vec<(&str, u32)>) {
        for (cell_ref, idx) in items {
            let f = format!(
                "=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,{},FALSE))",
                idx
            );
            self.set_formula(cell_ref, &f);
        }
    }

    pub fn write_to_sheet(&self, ws: &mut Worksheet) -> Result<(), XlsxError> {
        // 1. Column dimensions
        for (col, width) in &self.col_widths {
            ws.set_column_width(*col, *width)?;
        }
        for (start, end) in &self.grouped_cols {
            for c in *start..=*end {
                ws.set_column_hidden(c)?;
            }
        }
        for col in &self.hidden_cols {
            ws.set_column_hidden(*col)?;
        }

        // 2. Row dimensions
        for (row, height) in &self.row_heights {
            ws.set_row_height(*row, *height)?;
        }

        // 3. Merges
        let mut processed_cells = std::collections::HashSet::new();

        // Map top-left of merge to the merge range
        let mut merge_map = HashMap::new();
        for (r1, c1, r2, c2) in &self.merges {
            merge_map.insert((*r1, *c1), (*r2, *c2));
            for r in *r1..=*r2 {
                for c in *c1..=*c2 {
                    if r != *r1 || c != *c1 {
                        processed_cells.insert((r, c));
                    }
                }
            }
        }

        let mut all_coords: Vec<(u32, u16)> = self.cells.keys().cloned().collect();
        // Add merge start coords if not present
        for (r, c) in merge_map.keys() {
            if !self.cells.contains_key(&(*r, *c)) {
                all_coords.push((*r, *c));
            }
        }
        all_coords.sort();
        all_coords.dedup();

        for (row, col) in all_coords {
            if processed_cells.contains(&(row, col)) {
                continue;
            }

            let cell_default = VirtualCell::default();
            let cell = self.cells.get(&(row, col)).unwrap_or(&cell_default);

            let fmt = create_format(&cell.style);

            if let Some((r2, c2)) = merge_map.get(&(row, col)) {
                // Merge range
                match &cell.value {
                    CellValue::String(s) => ws.merge_range(row, col, *r2, *c2, s.as_str(), &fmt)?,
                    CellValue::Number(n) => {
                        ws.merge_range(row, col, *r2, *c2, "", &fmt)?;
                        ws.write_number_with_format(row, col, *n, &fmt)?
                    }
                    CellValue::Formula(f) => {
                        ws.merge_range(row, col, *r2, *c2, "", &fmt)?;
                        ws.write_formula_with_format(row, col, f.as_str(), &fmt)?
                    }
                    CellValue::Empty => ws.merge_range(row, col, *r2, *c2, "", &fmt)?,
                };
            } else {
                // Single cell
                match &cell.value {
                    CellValue::String(s) => {
                        ws.write_string_with_format(row, col, s.as_str(), &fmt)?
                    }
                    CellValue::Number(n) => ws.write_number_with_format(row, col, *n, &fmt)?,
                    CellValue::Formula(f) => {
                        ws.write_formula_with_format(row, col, f.as_str(), &fmt)?
                    }
                    CellValue::Empty => ws.write_blank(row, col, &fmt)?,
                };
            }
        }

        Ok(())
    }
}

fn create_format(style: &CellStyle) -> Format {
    let mut fmt = Format::new();

    // Set default font
    fmt = fmt.set_font_name(style.font_name.as_deref().unwrap_or("Arial"));
    fmt = fmt.set_font_size(style.font_size.unwrap_or(10.0));

    if let Some(true) = style.bold {
        fmt = fmt.set_bold();
    }
    if let Some(true) = style.italic {
        fmt = fmt.set_italic();
    }
    if let Some(u) = style.underline {
        fmt = fmt.set_underline(u);
    }
    if let Some(c) = style.font_color {
        fmt = fmt.set_font_color(c);
    }

    if let Some(c) = style.fill_color {
        fmt = fmt.set_background_color(c);
    }
    if let Some(p) = style.fill_pattern {
        fmt = fmt.set_pattern(p);
    }

    if let Some(h) = style.align_h {
        fmt = fmt.set_align(h);
    }
    if let Some(v) = style.align_v {
        fmt = fmt.set_align(v);
    }
    if let Some(true) = style.wrap_text {
        fmt = fmt.set_text_wrap();
    }

    if let Some(b) = style.border_left {
        fmt = fmt.set_border_left(b);
    }
    if let Some(b) = style.border_right {
        fmt = fmt.set_border_right(b);
    }
    if let Some(b) = style.border_top {
        fmt = fmt.set_border_top(b);
    }
    if let Some(b) = style.border_bottom {
        fmt = fmt.set_border_bottom(b);
    }

    if let Some(ref f) = style.num_format {
        fmt = fmt.set_num_format(f);
    }

    fmt
}

// Helpers
fn parse_range(range: &str) -> Option<(u32, u16, u32, u16)> {
    let re = Regex::new(r"^([A-Z]+)([0-9]+)(?::([A-Z]+)([0-9]+))?$").ok()?;
    let caps = re.captures(range)?;

    let c1_str = caps.get(1)?.as_str();
    let r1_str = caps.get(2)?.as_str();

    let c1 = col_str_to_index(c1_str)?;
    let r1 = r1_str.parse::<u32>().ok()? - 1; // 0-based

    if let (Some(c2_match), Some(r2_match)) = (caps.get(3), caps.get(4)) {
        let c2 = col_str_to_index(c2_match.as_str())?;
        let r2 = r2_match.as_str().parse::<u32>().ok()? - 1;
        Some((r1, c1, r2, c2))
    } else {
        Some((r1, c1, r1, c1))
    }
}

fn col_str_to_index(col: &str) -> Option<u16> {
    let mut num = 0;
    for c in col.chars() {
        if c < 'A' || c > 'Z' {
            return None;
        }
        num = num * 26 + (c as u16 - 'A' as u16 + 1);
    }
    if num == 0 {
        None
    } else {
        Some(num - 1)
    }
}

// ---------------------------------------------------------------------------
// Main Logic
// ---------------------------------------------------------------------------

pub fn recreate_sheet(workbook: &mut Workbook, sheet_name: &str) -> Result<(), XlsxError> {
    let mut model = SheetModel::new();

    // Constants
    let fill_input = Color::RGB(0xFFFAE5); // FFFFFAE5
    let fill_value = Color::RGB(0xF2F2F2);
    let fill_summary = Color::RGB(0xD9D9D9);
    let fl_orange = Color::RGB(0xFFCC00);
    // fl_light_yellow = fill_input

    // --- Column Dimensions ---
    model.set_col_width(0, 3.285); // A
    model.set_col_width(1, 4.57); // B
    model.set_col_width(2, 55.85); // C
    model.set_col_width(3, 20.0); // D
    model.set_col_width(4, 18.57); // E
    model.set_col_width(5, 18.71); // F
    model.set_col_width(6, 11.71); // G
    model.set_col_width(7, 18.71); // H
    model.set_col_width(8, 4.0); // I
    model.set_col_width(9, 4.0); // J
    model.set_col_width(10, 19.14); // K
    model.set_col_width(11, 10.85); // L
    model.set_col_width(12, 15.0); // M
    model.set_col_width(13, 15.0); // N
    model.set_col_width(14, 15.0); // O
    model.set_col_width(15, 5.0); // P
    model.set_col_width(16, 4.0); // Q
    model.set_col_width(17, 19.14); // R
    model.set_col_width(18, 10.85); // S
    model.set_col_width(19, 15.0); // T
    model.set_col_width(20, 15.0); // U
    model.set_col_width(21, 15.0); // V
    model.set_col_width(22, 4.0); // W
    model.set_col_width(23, 36.71); // X
    model.set_col_width(24, 10.85); // Y

    model.group_columns(16, 21); // Q-V

    // --- Row Dimensions ---
    model.set_row_height(3, 12.75); // Row 4 (0-based 3)
    model.set_row_height(4, 12.0);
    model.set_row_height(10, 13.5);
    model.set_row_height(11, 12.6);
    model.set_row_height(12, 12.6);
    model.set_row_height(13, 12.6);
    model.set_row_height(15, 12.75);
    model.set_row_height(16, 12.75);
    model.set_row_height(17, 12.75);
    model.set_row_height(18, 12.75);
    model.set_row_height(19, 12.75);
    model.set_row_height(20, 13.5);
    model.set_row_height(22, 13.5);
    model.set_row_height(23, 12.6);
    model.set_row_height(24, 12.6);
    model.set_row_height(25, 12.6);
    model.set_row_height(26, 12.6);

    // --- Merged Cells ---
    let merges = vec![
        "B16:C16", "E23:E26", "B25:C25", "B3:C3", "H23:H26", "J2:O3", "B6:C7", "G8:H8", "J4:O4",
        "D11:D14", "B18:C18", "G23:G26", "B12:C12", "B2:C2", "G9:H9", "H11:H14", "B17:C17",
        "B1:C1", "B8:C8", "F23:F26", "B13:C13", "J11:K11", "B19:C19", "D6:H7", "E11:E14", "B9:C9",
        "B24:C24", "D23:D26", "B15:C15", "Q11:R11", "B5:C5", "B20:C20", "F11:F14", "G11:G14",
    ];
    for m in merges {
        model.merge_cells(m);
    }

    // --- Alignments ---
    model.apply_alignment(
        &[
            "E8:E9", "G8:G9", "J4", "J11", "Q11", "B15:B20", "K14:K31", "R14:R31", "B182",
        ],
        Some(FormatAlign::Left),
        None,
        None,
    );

    model.apply_alignment(
        &["D15:G20", "J14:J31", "Q14:Q31"],
        Some(FormatAlign::Right),
        None,
        None,
    );

    model.apply_alignment(
        &["E2:E3", "L13:O13", "S13:V13"],
        Some(FormatAlign::Center),
        None,
        None,
    );

    model.apply_alignment(
        &["J2", "D6", "H15:H20"],
        Some(FormatAlign::Left),
        None,
        Some(true),
    );

    model.apply_alignment(
        &["B12:C14", "B24:B25"],
        Some(FormatAlign::Center),
        Some(FormatAlign::VerticalCenter),
        None,
    );

    model.apply_alignment(
        &["B1:C9", "D2:D3"],
        Some(FormatAlign::Left),
        Some(FormatAlign::VerticalCenter),
        None,
    );

    model.apply_alignment(
        &["B11:H14", "B23:H26"],
        Some(FormatAlign::Center),
        Some(FormatAlign::VerticalCenter),
        Some(true),
    );

    // --- Number Formats ---
    model.apply_number_format(&["D15:F20", "M14:N31", "T14:U31"], "#,##0.00");
    model.apply_number_format(&["E8:E9", "G8:G9", "L14:L31", "S14:S31"], "mm-dd-yy");
    model.apply_number_format(&["O14:O31", "V14:V31"], "0.0000");
    model.apply_number_format(&["G16:G20"], "0%");

    // --- Header ---
    model.set_formula(
        "D2",
        "=IF($E$2=\"\",\"Chose your language\",VLOOKUP($E$2,Sprachversionen!$B:$BN,27,FALSE))",
    );
    model.set_string("E2", "deutsch");
    model.set_formula(
        "J4",
        "=HYPERLINK(VLOOKUP($E$2,Sprachversionen!$B:$BN,62,FALSE))",
    );
    model.set_formula("D9", "=D8");
    model.set_formula("F9", "=F8");

    // --- Right Panel ---
    model.set_formula("J11", "=B18");
    model.set_formula("Q11", "=B18");
    model.set_string("M13", "Euro");
    model.set_formula(
        "N13",
        "=IF(E3=\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,28,FALSE),E3)",
    );
    model.set_string("T13", "Euro");
    model.set_formula(
        "U13",
        "=IF(E3=\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,28,FALSE),E3)",
    );

    for row in 14..=31 {
        let num = row - 13;
        model.set_string(&format!("J{}", row), &format!("{}. ", num));
        model.set_string(&format!("Q{}", row), &format!("{}. ", num + 18));
    }

    // Ratio Formulas
    model.apply_ratio_formulas(vec![
        ("O14:O31", "N", "M", 31, ""),
        ("V14:V31", "U", "T", 31, ""),
        ("G16:G20", "F", "D", 1001, "0"),
    ]);

    // D Column Zeros
    for row in 16..=19 {
        model.set_number(&format!("D{}", row), 0.0);
    }

    model.set_formula("D20", "=SUMPRODUCT(ROUND(D15:D19, 2))");
    model.set_formula("E20", "=SUMPRODUCT(ROUND(E15:E19, 2))");
    model.set_formula("F20", "=SUMPRODUCT(ROUND(F15:F19, 2))");
    model.set_formula(
        "G20",
        "=IFERROR(INDEX($F$1:$F$1001,ROW())/INDEX($D$1:$D$1001,ROW()),0)",
    );
    model.set_formula("D23", "=D11");
    model.set_formula("H23", "=H11");
    model.set_formula("B25", "=B13");

    // --- Fonts ---
    // Hyperlink Blue
    model.apply_font(&["J4:J5"], None, Some(Color::RGB(0x0563C1)), None, None);

    // Bold Black
    let bold_ranges = &[
        "B1:C3", "B12:C12", "B15:B20", "B24:C24", "E11:E14", "E23:E26", "J1", "J11:K11", "L13:O13",
        "Q11:R11", "S13:V13",
    ];
    model.apply_font(
        bold_ranges,
        Some(true),
        Some(Color::RGB(0x000000)),
        None,
        None,
    );

    // White
    model.apply_font(&["B2"], None, Some(Color::RGB(0xFFFFFF)), None, None);

    // --- Fills ---
    model.apply_fill(
        &[
            "E15:F19", "E8:E9", "G8:G9", "D5:D6", "H15:H19", "J9", "L14:N31", "S14:U31",
        ],
        fill_input,
    );
    model.apply_fill(&["E2:E3", "J7"], fl_orange);
    model.apply_fill(&["D15:D19", "J8"], fill_value);
    model.apply_fill(&["B20:H20"], fill_summary);

    // --- Borders ---

    // D5, D6:H9 - Bottom only, thin
    model.apply_grid_border(
        &["D5", "D6:H9"],
        &["bottom"],
        "horizontal",
        FormatBorder::Thin,
        Color::Black,
    );

    // E2:E3 - Bottom dashed
    model.apply_grid_border(
        &["E2:E3"],
        &["bottom"],
        "horizontal",
        FormatBorder::Dashed,
        Color::Black,
    );

    // J7:J9 - All Dotted
    model.apply_grid_border(
        &["J7:J9"],
        &["all"],
        "horizontal",
        FormatBorder::Dotted,
        Color::Black,
    );

    // D11:H20, L14:O31, S14:V31 - All Thin, Inner Both
    model.apply_grid_border(
        &["D11:H20", "L14:O31", "S14:V31"],
        &["all"],
        "both",
        FormatBorder::Thin,
        Color::Black,
    );

    // B15:C19, ... - All Thin, Inner None
    model.apply_grid_border(
        &["B15:C19", "J11:O31", "Q11:V31"],
        &["all"],
        "none",
        FormatBorder::Thin,
        Color::Black,
    );

    // B11:H20, J1:O4 - All Medium, Inner None
    model.apply_grid_border(
        &["B11:H20", "J1:O4"],
        &["all"],
        "none",
        FormatBorder::Medium,
        Color::Black,
    );

    // B23:H26 - Left, Top, Right Medium. Inner None.
    model.apply_grid_border(
        &["B23:H26"],
        &["left", "top", "right"],
        "none",
        FormatBorder::Medium,
        Color::Black,
    );

    // D23:H26 - Left, Bottom Thin. Inner Vertical Thin.
    model.apply_grid_border(
        &["D23:H26"],
        &["left", "bottom"],
        "vertical",
        FormatBorder::Thin,
        Color::Black,
    );

    // B11:C14 - Right Thin
    model.apply_grid_border(
        &["B11:C14"],
        &["right"],
        "none",
        FormatBorder::Thin,
        Color::Black,
    );

    // B23:C26 - Right Thin
    model.apply_grid_border(
        &["B23:C26"],
        &["right"],
        "none",
        FormatBorder::Thin,
        Color::Black,
    );

    // --- VLOOKUPs ---
    model.apply_vlookup_formulas(vec![
        ("B1", 2),
        ("J1", 60),
        ("B3", 3),
        ("D3", 28),
        ("J2", 61),
        ("B5", 4),
        ("B6", 26),
        ("B8", 5),
        ("D8", 7),
        ("F8", 8),
        ("K8", 59),
        ("B9", 6),
        ("D11", 11),
        ("E11", 12),
        ("F11", 13),
        ("G11", 14),
        ("H11", 15),
        ("B12", 9),
        ("L13", 22),
        ("O13", 58),
        ("S13", 22),
        ("V13", 58),
        ("B13", 10),
        ("B15", 16),
        ("B16", 17),
        ("B17", 18),
        ("B18", 19),
        ("B19", 20),
        ("B20", 21),
    ]);

    for i in 0..18 {
        let row = 14 + i;
        let f = format!("=IF($E$2=\"\",\"\",VLOOKUP($E$2,Sprachversionen!$B:$BN,23,FALSE))");
        model.set_formula(&format!("K{}", row), &f);
        model.set_formula(&format!("R{}", row), &f);
    }

    model.apply_vlookup_formulas(vec![("E23", 25), ("F23", 55), ("G23", 56), ("B24", 24)]);

    // Final Write
    let ws = workbook.add_worksheet().set_name(sheet_name)?;
    ws.set_zoom(85);
    ws.set_screen_gridlines(false);

    model.write_to_sheet(ws)?;

    Ok(())
}
