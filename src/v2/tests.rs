#[cfg(test)]
mod tests {
    use crate::v2::Finanzbericht::header::write_header;
    use crate::v2::Finanzbericht::sheet_setup::sheet_setup;
    use crate::v2::Finanzbericht::styles::ReportStyles;
    use crate::v2::Sprachversion::builder::build_sheet as build_trans_sheet;
    use rust_xlsxwriter::Workbook;

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

        sheet_setup(ws).expect("Failed to setup sheet");

        // 3. Prepare Styles
        let styles = ReportStyles::new();

        // 4. Write Header
        let suffix = "_de";
        let lang_val = "deutsch";
        write_header(ws, &styles, suffix, lang_val).expect("Failed to write header");

        // 5. Save to file for inspection
        let path = "src/v2/tests/header_test.xlsx";
        workbook
            .save(path)
            .expect("Failed to save workbook to file");
    }
}
