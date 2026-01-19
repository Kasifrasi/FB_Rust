#[cfg(test)]
mod tests_right_panel {
    use crate::v2::report::formats::ReportStyles;
    use crate::v2::report::layout::setup_sheet;
    use crate::v2::report::values::ReportValues;
    use crate::v2::report::writer::write_report_v2;
    use crate::v2::Sprachversion::builder::build_sheet as build_trans_sheet;
    use rust_xlsxwriter::{Format, Workbook};

    /// Test mit vollständigen Right Panel Daten
    #[test]
    fn test_right_panel_with_complete_data() {
        let mut workbook = Workbook::new();

        let sheet_name = "Finanzbericht";
        let _ = workbook.add_worksheet().set_name(sheet_name).unwrap();

        build_trans_sheet(&mut workbook).expect("Failed to build translation sheet");

        let ws = workbook
            .worksheet_from_name(sheet_name)
            .expect("Sheet not found");

        let unlocked = Format::new()
            .set_font_name("Arial")
            .set_font_size(10.0)
            .set_unlocked();

        for col in 0..1000u16 {
            ws.set_column_format(col, &unlocked)
                .expect("Failed to set column format");
        }

        setup_sheet(ws).expect("Failed to setup sheet");

        let styles = ReportStyles::new();

        let mut values = ReportValues::new()
            .with_language("deutsch")
            .with_currency("EUR")
            .with_project_number("TEST-RP-001")
            .with_project_title("Right Panel Test")
            .with_exchange_rate(1.1);

        use crate::v2::report::cells::{RightPanelInputCell, TableInputCell};

        for i in 0..5u8 {
            values
                .set(TableInputCell::ApprovedBudget(i), 1000.0 + i as f64 * 500.0)
                .set(
                    TableInputCell::IncomeReportPeriod(i),
                    100.0 + i as f64 * 200.0,
                )
                .set(TableInputCell::IncomeTotal(i), 500.0 + i as f64 * 300.0);
        }

        // Right Panel Daten - alle 18 Zeilen mit Datums- und Zahlenwerten
        for i in 0..18u8 {
            let num = i + 1;
            let amount_eur = 100.0 + i as f64 * 250.0;
            let amount_local = 110.0 + i as f64 * 275.0;

            let day = (i % 28) + 1;
            let month = 3 + (i / 10);
            let date_str = format!("2024-{:02}-{:02}", month, day);

            values
                .set(RightPanelInputCell::LeftNumber(i), num.to_string())
                .set(RightPanelInputCell::LeftDate(i), date_str.clone())
                .set(RightPanelInputCell::LeftAmount1(i), amount_eur)
                .set(RightPanelInputCell::LeftAmount2(i), amount_local)
                .set(RightPanelInputCell::RightNumber(i), (num + 100).to_string())
                .set(RightPanelInputCell::RightDate(i), date_str)
                .set(RightPanelInputCell::RightAmount1(i), amount_eur * 1.5)
                .set(RightPanelInputCell::RightAmount2(i), amount_local * 1.5);
        }

        let suffix = "_de";
        write_report_v2(ws, &styles, suffix, &values).expect("Failed to write report");

        ws.protect();

        let path = "src/v2/tests/right_panel_complete_data.xlsx";
        workbook.save(path).expect("Failed to save workbook");

        println!("✅ Right Panel Test erfolgreich!");
        println!("   Datei: {}", path);
        println!("   - 18 Zeilen Daten in L14-L31 (Datumswerte)");
        println!("   - 18 Zeilen Zahlen in M14-M31 (Euro)");
        println!("   - 18 Zeilen Zahlen in N14-N31 (Lokal)");
        println!("   - K14-K31: Nummern 1-18");
        println!("   - R14-R31: Nummern 101-118");
        println!("   - S14-S31: Datumswerte (wie Spalte L)");
        println!("   - T14-T31: 1.5x der Werte aus M");
        println!("   - U14-U31: 1.5x der Werte aus N");
    }

    /// Test mit variierenden Datumswerten und Beträgen
    #[test]
    fn test_right_panel_realistic_data() {
        let mut workbook = Workbook::new();

        let sheet_name = "Finanzbericht";
        let _ = workbook.add_worksheet().set_name(sheet_name).unwrap();

        build_trans_sheet(&mut workbook).expect("Failed to build translation sheet");

        let ws = workbook
            .worksheet_from_name(sheet_name)
            .expect("Sheet not found");

        let unlocked = Format::new()
            .set_font_name("Arial")
            .set_font_size(10.0)
            .set_unlocked();

        for col in 0..1000u16 {
            ws.set_column_format(col, &unlocked).ok();
        }

        setup_sheet(ws).expect("Failed to setup sheet");

        let styles = ReportStyles::new();

        let mut values = ReportValues::new()
            .with_language("deutsch")
            .with_currency("EUR")
            .with_project_number("TEST-RP-002")
            .with_project_title("Realistic Right Panel Data")
            .with_exchange_rate(1.15);

        use crate::v2::report::cells::RightPanelInputCell;

        // Realistische Daten mit variierenden Beträgen
        let realistic_amounts = vec![
            250.50, 500.75, 1200.00, 450.25, 850.00, 650.50, 920.75, 1100.00, 380.25, 760.00,
            1050.50, 670.75, 890.00, 520.25, 1300.50, 440.00, 680.75, 810.00,
        ];

        for (idx, &amount) in realistic_amounts.iter().enumerate() {
            let i = idx as u8;
            let num = i + 1;
            let amount_local = amount * 1.15;

            let date_str = match i {
                0..=9 => format!("2024-03-{:02}", i + 1),
                10..=17 => format!("2024-04-{:02}", i - 9),
                _ => format!("2024-05-01"),
            };

            values
                .set(
                    RightPanelInputCell::LeftNumber(i),
                    format!("RGG-{:03}", 100 + num),
                )
                .set(RightPanelInputCell::LeftDate(i), date_str.clone())
                .set(RightPanelInputCell::LeftAmount1(i), amount)
                .set(RightPanelInputCell::LeftAmount2(i), amount_local)
                .set(
                    RightPanelInputCell::RightNumber(i),
                    format!("INV-{:03}", 200 + num),
                )
                .set(RightPanelInputCell::RightDate(i), date_str)
                .set(RightPanelInputCell::RightAmount1(i), amount * 1.25)
                .set(RightPanelInputCell::RightAmount2(i), amount_local * 1.25);
        }

        let suffix = "_de";
        write_report_v2(ws, &styles, suffix, &values).expect("Failed to write report");

        ws.protect();

        let path = "src/v2/tests/right_panel_realistic_data.xlsx";
        workbook.save(path).expect("Failed to save workbook");

        println!("✅ Realistic Right Panel Test erfolgreich!");
        println!("   Datei: {}", path);
        println!("   - Realistische Rechnungsnummern (RGG-101..118, INV-201..218)");
        println!("   - Datumswerte März-April 2024");
        println!("   - Variierte Beträge (250-1300 EUR)");
        println!("   - Lokale Währung berechnet (EUR * Wechselkurs 1.15)");
    }
}
