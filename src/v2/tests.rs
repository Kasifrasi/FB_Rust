#[cfg(test)]
mod tests {
    use crate::v2::Finanzbericht::header::write_header;
    use crate::v2::Finanzbericht::sheet_setup::sheet_setup;
    use crate::v2::Finanzbericht::styles::ReportStyles;
    use crate::v2::Sprachversion::builder::build_sheet as build_trans_sheet;
    use rust_xlsxwriter::{Format, Workbook};

    #[test]
    fn test_header_generation() {
        let mut workbook = Workbook::new();

        // 1. Add Target Sheet (Finanzbericht) - must be first to be on left
        let sheet_name = "Finanzbericht";
        let _ = workbook.add_worksheet().set_name(sheet_name).unwrap();

        // 2. Build Translation Sheet (Reference for VLOOKUPs)
        build_trans_sheet(&mut workbook).expect("Failed to build translation sheet");

        // 3. Get Target Sheet back
        let ws = workbook
            .worksheet_from_name(sheet_name)
            .expect("Sheet not found");

        // 4. Set column format to unlocked for 1000 columns
        // This makes all cells in these columns unlocked by default
        let unlocked = Format::new()
            .set_font_name("Arial")
            .set_font_size(10.0)
            .set_unlocked();

        for col in 0..1000u16 {
            ws.set_column_format(col, &unlocked)
                .expect("Failed to set column format");
        }

        // 5. Setup sheet (column widths, etc.)
        sheet_setup(ws).expect("Failed to setup sheet");

        // 6. Prepare Styles
        let styles = ReportStyles::new();

        // 7. Write Header
        let suffix = "_de";
        let lang_val = "deutsch";
        write_header(ws, &styles, suffix, lang_val).expect("Failed to write header");

        // 8. Protect worksheet
        // All cells are unlocked by set_row_format()
        // Formulas are locked by write_formulas() which uses fmt.get_locked()
        ws.protect();

        // 9. Save to file for inspection
        let path = "src/v2/tests/header_test.xlsx";
        workbook
            .save(path)
            .expect("Failed to save workbook to file");
    }
}
