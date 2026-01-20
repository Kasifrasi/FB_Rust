#[cfg(test)]
mod tests {
    use crate::v2::Finanzbericht::header::write_header;
    use crate::v2::Finanzbericht::sheet_setup::sheet_setup;
    use crate::v2::Finanzbericht::styles::ReportStyles;
    use crate::v2::Finanzbericht::values::ReportValues;
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

        // 7. Prepare Values with language set to "deutsch"
        let values = ReportValues::new().with_language("deutsch");

        // 8. Write Header
        let suffix = "_de";
        write_header(ws, &styles, suffix, &values).expect("Failed to write header");

        // 9. Protect worksheet
        // All cells are unlocked by set_row_format()
        // Formulas are locked by write_formulas() which uses fmt.get_locked()
        ws.protect();

        // 10. Save to file for inspection
        let path = "src/v2/tests/header_test.xlsx";
        workbook
            .save(path)
            .expect("Failed to save workbook to file");
    }

    /// Test mit neuem Registry-basierten Writer
    #[test]
    fn test_header_generation_v2() {
        use crate::v2::report::formats::ReportStyles;
        use crate::v2::report::layout::setup_sheet;
        use crate::v2::report::values::ReportValues;
        use crate::v2::report::writer::write_report_v2;

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
        write_report_v2(ws, &styles, suffix, &values).expect("Failed to write report v2");

        // 9. Protect worksheet
        ws.protect();

        // 10. Save to file
        let path = "src/v2/tests/header_test_v2.xlsx";
        workbook
            .save(path)
            .expect("Failed to save workbook to file");
    }

    /// Test mit numerischen Werten um Formelauswertung zu verifizieren
    #[test]
    fn test_formula_evaluation_with_numbers() {
        use crate::v2::report::definitions::build_registry;
        use crate::v2::report::formats::ReportStyles;
        use crate::v2::report::layout::setup_sheet;
        use crate::v2::report::registry::{ApiKey, CellAddr, CellKind, EvalContext};
        use crate::v2::report::values::{CellValue, ReportValues};
        use crate::v2::report::writer::write_report_v2;
        use std::collections::HashMap;

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
        write_report_v2(ws, &styles, suffix, &values).expect("Failed to write report");

        // 9. Protect worksheet
        ws.protect();

        // 10. Save to file
        let path = "src/v2/tests/formula_eval_test.xlsx";
        workbook
            .save(path)
            .expect("Failed to save workbook to file");

        // 11. Verifiziere Formel-Auswertung durch Registry-Evaluation
        let registry = build_registry().expect("Failed to build registry");

        // Evaluiere alle Zellen
        let mut computed: HashMap<CellAddr, CellValue> = HashMap::new();

        // API-Werte eintragen
        for addr in registry.api_cells() {
            if let CellKind::Api(api) = registry.get(*addr) {
                let value = match api.key {
                    ApiKey::Language => CellValue::Text("deutsch".to_string()),
                    ApiKey::Currency => CellValue::Text("EUR".to_string()),
                    ApiKey::ApprovedBudget(0) => CellValue::Number(1000.0),
                    ApiKey::ApprovedBudget(1) => CellValue::Number(2000.0),
                    ApiKey::ApprovedBudget(2) => CellValue::Number(3000.0),
                    ApiKey::ApprovedBudget(3) => CellValue::Number(4000.0),
                    ApiKey::ApprovedBudget(4) => CellValue::Number(5000.0),
                    ApiKey::IncomeReportPeriod(0) => CellValue::Number(100.0),
                    ApiKey::IncomeReportPeriod(1) => CellValue::Number(200.0),
                    ApiKey::IncomeReportPeriod(2) => CellValue::Number(300.0),
                    ApiKey::IncomeReportPeriod(3) => CellValue::Number(400.0),
                    ApiKey::IncomeReportPeriod(4) => CellValue::Number(500.0),
                    ApiKey::IncomeTotal(0) => CellValue::Number(500.0),
                    ApiKey::IncomeTotal(1) => CellValue::Number(1000.0),
                    ApiKey::IncomeTotal(2) => CellValue::Number(1500.0),
                    ApiKey::IncomeTotal(3) => CellValue::Number(2000.0),
                    ApiKey::IncomeTotal(4) => CellValue::Number(2500.0),
                    _ => CellValue::Empty,
                };
                computed.insert(*addr, value);
            }
        }

        // Formeln evaluieren
        let mut formula_addrs: Vec<CellAddr> = registry.formula_cells().iter().copied().collect();
        formula_addrs.sort();

        for addr in formula_addrs {
            if let CellKind::Formula(f) = registry.get(addr) {
                let ctx = EvalContext {
                    computed: &computed,
                    api_values: &values,
                };
                let result = (f.eval)(&ctx);
                computed.insert(addr, result);
            }
        }

        // Verifiziere einige berechnete Werte
        // D20 sollte SUMPRODUCT(D15:D19) = 1000+2000+3000+4000+5000 = 15000 sein
        let d20 = computed.get(&CellAddr::new(19, 3));
        assert!(d20.is_some(), "D20 should be computed");
        if let Some(CellValue::Number(n)) = d20 {
            assert_eq!(*n, 15000.0, "D20 should be 15000 (sum of budgets)");
        }

        // F20 sollte SUMPRODUCT(F15:F19) = 500+1000+1500+2000+2500 = 7500 sein
        let f20 = computed.get(&CellAddr::new(19, 5));
        assert!(f20.is_some(), "F20 should be computed");
        if let Some(CellValue::Number(n)) = f20 {
            assert_eq!(*n, 7500.0, "F20 should be 7500 (sum of income totals)");
        }

        // G20 sollte F20/D20 = 7500/15000 = 0.5 sein
        let g20 = computed.get(&CellAddr::new(19, 6));
        assert!(g20.is_some(), "G20 should be computed");
        if let Some(CellValue::Number(n)) = g20 {
            assert!(
                (n - 0.5).abs() < 0.001,
                "G20 should be 0.5 (7500/15000), got {}",
                n
            );
        }

        println!("Formula evaluation test passed!");
        println!("D20 (sum of budgets): {:?}", d20);
        println!("F20 (sum of income totals): {:?}", f20);
        println!("G20 (F20/D20): {:?}", g20);
    }
}
