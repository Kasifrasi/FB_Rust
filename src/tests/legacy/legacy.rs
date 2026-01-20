#[cfg(test)]
mod tests {
    use crate::lang::builder::build_sheet as build_trans_sheet;
    use crate::report::formats::ReportStyles;
    use crate::report::layout::setup_sheet;
    use crate::report::protection::HiddenRanges;
    use crate::report::values::ReportValues;
    use crate::report::writer::write_report;
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
        setup_sheet(ws).expect("Failed to setup sheet");

        // 6. Prepare Styles
        let styles = ReportStyles::new();

        // 7. Prepare Values with language set to "deutsch"
        let values = ReportValues::new().with_language("deutsch");

        // 8. Write Header
        let suffix = "_de";
        write_report(ws, &styles, suffix, &values).expect("Failed to write header");

        // 9. Spalten Q:V verstecken und Sheet schützen
        let hidden = HiddenRanges::preset_hide_qv();
        for range in hidden.column_ranges() {
            for col in range.start..=range.end {
                ws.set_column_hidden(col as u16).ok();
            }
        }
        ws.protect();

        // 10. Save to file for inspection
        let path = "tests/output/header_test.xlsx";
        workbook
            .save(path)
            .expect("Failed to save workbook to file");
    }

    /// Test mit neuem Registry-basierten Writer
    #[test]
    fn test_header_generation_v2() {
        use crate::report::formats::ReportStyles;
        use crate::report::layout::setup_sheet;
        use crate::report::values::ReportValues;
        use crate::report::writer::write_report;

        let mut workbook = Workbook::new();

        // 1. Add Target Sheet
        let sheet_name = "Finanzbericht";
        let _ = workbook.add_worksheet().set_name(sheet_name).unwrap();

        // 2. Build Translation Sheet
        build_trans_sheet(&mut workbook).expect("Failed to build translation sheet");

        // 3. Get Target Sheet back
        let ws = workbook
            .worksheet_from_name(sheet_name)
            .expect("Sheet not found");

        // 4. Set column format to unlocked
        let unlocked = Format::new()
            .set_font_name("Arial")
            .set_font_size(10.0)
            .set_unlocked();

        for col in 0..1000u16 {
            ws.set_column_format(col, &unlocked)
                .expect("Failed to set column format");
        }

        // 5. Setup sheet
        setup_sheet(ws).expect("Failed to setup sheet");

        // 6. Prepare Styles
        let styles = ReportStyles::new();

        // 7. Prepare Values
        let values = ReportValues::new()
            .with_language("deutsch")
            .with_currency("EUR")
            .with_project_number("12345")
            .with_project_title("Test Projekt");

        // 8. Write Report mit V2 Writer
        let suffix = "_de";
        write_report(ws, &styles, suffix, &values).expect("Failed to write report v2");

        // 9. Spalten Q:V verstecken und Sheet schützen
        let hidden = HiddenRanges::preset_hide_qv();
        for range in hidden.column_ranges() {
            for col in range.start..=range.end {
                ws.set_column_hidden(col as u16).ok();
            }
        }
        ws.protect();

        // 10. Save to file
        let path = "tests/output/header_test_v2.xlsx";
        workbook
            .save(path)
            .expect("Failed to save workbook to file");
    }

    /// Test mit numerischen Werten um Formelauswertung zu verifizieren
    #[test]
    fn test_formula_evaluation_with_numbers() {
        use crate::report::definitions::build_registry;
        use crate::report::formats::ReportStyles;
        use crate::report::layout::setup_sheet;
        use crate::report::registry::{ApiKey, CellAddr};
        use crate::report::values::ReportValues;
        use crate::report::writer::write_report;

        let mut workbook = Workbook::new();

        // 1. Add Target Sheet
        let sheet_name = "Finanzbericht";
        let _ = workbook.add_worksheet().set_name(sheet_name).unwrap();

        // 2. Build Translation Sheet
        build_trans_sheet(&mut workbook).expect("Failed to build translation sheet");

        // 3. Get Target Sheet back
        let ws = workbook
            .worksheet_from_name(sheet_name)
            .expect("Sheet not found");

        // 4. Set column format to unlocked
        let unlocked = Format::new()
            .set_font_name("Arial")
            .set_font_size(10.0)
            .set_unlocked();

        for col in 0..1000u16 {
            ws.set_column_format(col, &unlocked)
                .expect("Failed to set column format");
        }

        // 5. Setup sheet
        setup_sheet(ws).expect("Failed to setup sheet");

        // 6. Prepare Styles
        let styles = ReportStyles::new();

        // 7. Prepare Values mit numerischen Testwerten
        let mut values = ReportValues::new()
            .with_language("deutsch")
            .with_currency("EUR")
            .with_project_number("TEST-12345")
            .with_project_title("Formel-Test Projekt");

        // Tabellen-Werte setzen (D15:F19) - nutze direkt ApiKey

        // D15-D19: Bewilligtes Budget (1000, 2000, 3000, 4000, 5000)
        values.set(ApiKey::ApprovedBudget(0), 1000.0);
        values.set(ApiKey::ApprovedBudget(1), 2000.0);
        values.set(ApiKey::ApprovedBudget(2), 3000.0);
        values.set(ApiKey::ApprovedBudget(3), 4000.0);
        values.set(ApiKey::ApprovedBudget(4), 5000.0);

        // E15-E19: Einnahmen Berichtszeitraum (100, 200, 300, 400, 500)
        values.set(ApiKey::IncomeReportPeriod(0), 100.0);
        values.set(ApiKey::IncomeReportPeriod(1), 200.0);
        values.set(ApiKey::IncomeReportPeriod(2), 300.0);
        values.set(ApiKey::IncomeReportPeriod(3), 400.0);
        values.set(ApiKey::IncomeReportPeriod(4), 500.0);

        // F15-F19: Einnahmen gesamt (500, 1000, 1500, 2000, 2500)
        values.set(ApiKey::IncomeTotal(0), 500.0);
        values.set(ApiKey::IncomeTotal(1), 1000.0);
        values.set(ApiKey::IncomeTotal(2), 1500.0);
        values.set(ApiKey::IncomeTotal(3), 2000.0);
        values.set(ApiKey::IncomeTotal(4), 2500.0);

        // 8. Write Report
        let suffix = "_de";
        write_report(ws, &styles, suffix, &values).expect("Failed to write report");

        // 9. Spalten Q:V verstecken und Sheet schützen
        let hidden = HiddenRanges::preset_hide_qv();
        for range in hidden.column_ranges() {
            for col in range.start..=range.end {
                ws.set_column_hidden(col as u16).ok();
            }
        }
        ws.protect();

        // 10. Save to file
        let path = "tests/output/formula_eval_test.xlsx";
        workbook
            .save(path)
            .expect("Failed to save workbook to file");

        // 11. Verifiziere dass Registry API-Zellen UND Formeln enthält
        let registry = build_registry().expect("Failed to build registry");

        // API-Zellen sollten registriert sein
        assert!(
            registry.api_cells().len() > 50,
            "API cells should be registered"
        );

        // Statische Formeln (wie D20, F20, G20) sind in der Registry registriert
        assert!(
            registry.is_formula(CellAddr::new(19, 3)),
            "D20 should be in registry"
        );
        assert!(
            registry.is_formula(CellAddr::new(19, 5)),
            "F20 should be in registry"
        );
        assert!(
            registry.is_formula(CellAddr::new(19, 6)),
            "G20 should be in registry"
        );

        println!("Formula evaluation test passed!");
        println!("All formulas are registered in the Registry");
        println!("write_cells_from_registry() writes them with correct formats from FormatMatrix");
    }
}
