#[cfg(test)]
mod tests_right_panel {
    use crate::lang::builder::build_sheet as build_trans_sheet;
    use crate::report::api::ApiKey;
    use crate::report::formats::ReportStyles;
    use crate::report::layout::{hide_columns_qv, setup_sheet};
    use crate::report::values::ReportValues;
    use crate::report::writer::write_report;
    use rust_xlsxwriter::{Format, Workbook};

    /// Test mit vollständigen Right Panel Daten
    ///
    /// NUR diese Zellen sind API-Eingabefelder:
    /// - L14:N31 (Datum, Euro, Lokal - linke Seite)
    /// - S14:U31 (Datum, Euro, Lokal - rechte Seite)
    ///
    /// K14:K31 und R14:R31 sind FORMELN und dürfen NICHT von API befüllt werden!
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
            .with_project_start("2024-01-01")
            .with_project_end("2024-12-31")
            .with_report_start("2024-01-01")
            .with_report_end("2024-03-31");

        // Tabellendaten (D15-F19, H15-H19)
        for i in 0..5u8 {
            values
                .set(ApiKey::ApprovedBudget(i), 1000.0 + i as f64 * 500.0)
                .set(ApiKey::IncomeReportPeriod(i), 100.0 + i as f64 * 200.0)
                .set(ApiKey::IncomeTotal(i), 500.0 + i as f64 * 300.0);
        }

        // Right Panel Daten - NUR L, M, N (links) und S, T, U (rechts)
        // K14:K31 und R14:R31 sind FORMELN, nicht befüllen!
        for i in 0..18u8 {
            let amount_eur = 100.0 + i as f64 * 250.0;
            let amount_local = 110.0 + i as f64 * 275.0;

            let day = (i % 28) + 1;
            let month = 3 + (i / 10);
            let date_str = format!("2024-{:02}-{:02}", month, day);

            values
                // Linke Seite: L, M, N (Datum, Euro, Lokal)
                .set(ApiKey::LeftDate(i), date_str.clone())
                .set(ApiKey::LeftAmountEuro(i), amount_eur)
                .set(ApiKey::LeftAmountLocal(i), amount_local)
                // Rechte Seite: S, T, U (Datum, Euro, Lokal)
                .set(ApiKey::RightDate(i), date_str)
                .set(ApiKey::RightAmountEuro(i), amount_eur * 1.5)
                .set(ApiKey::RightAmountLocal(i), amount_local * 1.5);
        }

        let suffix = "_de";
        write_report(ws, &styles, suffix, &values).expect("Failed to write report");

        // Spalten Q:V verstecken und Sheet schützen
        hide_columns_qv(ws).expect("Failed to hide columns");
        ws.protect();

        let path = "tests/output/right_panel_complete_data.xlsx";
        workbook.save(path).expect("Failed to save workbook");

        println!("Right Panel Test erfolgreich!");
        println!("   Datei: {}", path);
        println!("   - L14-L31: Datumswerte");
        println!("   - M14-M31: Euro-Betraege");
        println!("   - N14-N31: Lokal-Betraege");
        println!("   - S14-S31: Datumswerte");
        println!("   - T14-T31: Euro-Betraege (1.5x)");
        println!("   - U14-U31: Lokal-Betraege (1.5x)");
        println!("   - K14:K31, R14:R31 bleiben als Formeln erhalten!");
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
            .with_project_start("2024-01-01")
            .with_project_end("2024-12-31")
            .with_report_start("2024-01-01")
            .with_report_end("2024-03-31");

        // Realistische Daten mit variierenden Beträgen
        let realistic_amounts = vec![
            250.50, 500.75, 1200.00, 450.25, 850.00, 650.50, 920.75, 1100.00, 380.25, 760.00,
            1050.50, 670.75, 890.00, 520.25, 1300.50, 440.00, 680.75, 810.00,
        ];

        for (idx, &amount) in realistic_amounts.iter().enumerate() {
            let i = idx as u8;
            let amount_local = amount * 1.15;

            let date_str = match i {
                0..=9 => format!("2024-03-{:02}", i + 1),
                10..=17 => format!("2024-04-{:02}", i - 9),
                _ => format!("2024-05-01"),
            };

            values
                // Linke Seite: L, M, N (Datum, Euro, Lokal)
                .set(ApiKey::LeftDate(i), date_str.clone())
                .set(ApiKey::LeftAmountEuro(i), amount)
                .set(ApiKey::LeftAmountLocal(i), amount_local)
                // Rechte Seite: S, T, U (Datum, Euro, Lokal)
                .set(ApiKey::RightDate(i), date_str)
                .set(ApiKey::RightAmountEuro(i), amount * 1.25)
                .set(ApiKey::RightAmountLocal(i), amount_local * 1.25);
        }

        let suffix = "_de";
        write_report(ws, &styles, suffix, &values).expect("Failed to write report");

        // Spalten Q:V verstecken und Sheet schützen
        hide_columns_qv(ws).expect("Failed to hide columns");
        ws.protect();

        let path = "tests/output/right_panel_realistic_data.xlsx";
        workbook.save(path).expect("Failed to save workbook");

        println!("Realistic Right Panel Test erfolgreich!");
        println!("   Datei: {}", path);
        println!("   - Datumswerte Maerz-April 2024");
        println!("   - Variierte Betraege (250-1300 EUR)");
        println!("   - Lokale Waehrung berechnet (EUR * Wechselkurs 1.15)");
        println!("   - K14:K31, R14:R31 bleiben als Formeln erhalten!");
    }
}
