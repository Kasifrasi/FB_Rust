//! Zentrale Testdatei für den Finanzbericht-Generator
//!
//! Diese Datei enthält alle aktiven Tests und generiert Excel-Dateien zur Überprüfung.
//! Verwendet immer die neueste API.
//!
//! Generierte Dateien: tests/output/*.xlsx

#[cfg(test)]
mod report_tests {
    use crate::lang::builder::build_sheet as build_trans_sheet;
    use crate::report::api::ApiKey;
    use crate::report::formats::ReportStyles;
    use crate::report::layout::setup_sheet;
    use crate::report::protection::ReportOptions;
    use crate::report::values::ReportValues;
    use crate::report::writer::write_report_with_options;
    use crate::report::BodyConfig;
    use rust_xlsxwriter::{Format, Workbook};

    /// Hilfsfunktion: Erstellt eine Workbook mit Standard-Setup
    fn create_workbook_with_sheet(sheet_name: &str) -> Workbook {
        let mut workbook = Workbook::new();
        let _ = workbook.add_worksheet().set_name(sheet_name).unwrap();
        build_trans_sheet(&mut workbook).expect("Failed to build translation sheet");
        workbook
    }

    /// Hilfsfunktion: Setup für ein Worksheet
    fn setup_worksheet(workbook: &mut Workbook, sheet_name: &str) {
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
    }

    // ========================================================================
    // Test 1: Mittleres Projekt - Realistisches Szenario
    // ========================================================================
    #[test]
    fn test_medium_project() {
        let sheet_name = "Finanzbericht";
        let mut workbook = create_workbook_with_sheet(sheet_name);
        setup_worksheet(&mut workbook, sheet_name);

        let ws = workbook.worksheet_from_name(sheet_name).unwrap();
        let styles = ReportStyles::new();

        let mut values = ReportValues::new()
            .with_language("deutsch")
            .with_currency("EUR")
            .with_project_number("KMW-2024-001")
            .with_project_title("Mittelpunktanalyse von Forschungsmethoden")
            .with_project_start("01.01.2024")
            .with_project_end("31.12.2024")
            .with_report_start("01.01.2024")
            .with_report_end("30.06.2024");

        // Tabellen-Daten (Zeilen 15-19)
        let budgets = [50000.0, 30000.0, 15000.0, 8000.0, 5000.0];
        let incomes = [25000.0, 15000.0, 7500.0, 4000.0, 2500.0];
        for i in 0..5u8 {
            values
                .set(ApiKey::ApprovedBudget(i), budgets[i as usize])
                .set(ApiKey::IncomeReportPeriod(i), incomes[i as usize] * 1.1)
                .set(ApiKey::IncomeTotal(i), incomes[i as usize])
                .set(ApiKey::IncomeReason(i), format!("Finanzierung {}", i + 1));
        }

        // Right Panel (L14-N31, S14-U31)
        for i in 0..18u8 {
            let amount = 1000.0 + i as f64 * 200.0;
            let date = format!("2024-{:02}-{:02}", ((i / 6) % 6) + 1, ((i % 6) * 5) + 1);
            values
                .set(ApiKey::LeftDate(i), date.clone())
                .set(ApiKey::LeftAmountEuro(i), amount)
                .set(ApiKey::LeftAmountLocal(i), amount)
                .set(ApiKey::RightDate(i), date)
                .set(ApiKey::RightAmountEuro(i), amount * 1.2)
                .set(ApiKey::RightAmountLocal(i), amount * 1.2);
        }

        // Kategorie 1: Personal (10 Positionen)
        let personnel = [
            ("Projektleiter (100%)", 60000.0),
            ("Senior Wissenschaftler (100%)", 55000.0),
            ("Wissenschaftlicher Mitarbeiter 1 (80%)", 32000.0),
            ("Wissenschaftlicher Mitarbeiter 2 (80%)", 32000.0),
            ("Techniker/in (50%)", 18000.0),
            ("Sekretariat/Administration (30%)", 10000.0),
            ("Studentische Hilfskraft 1", 3600.0),
            ("Studentische Hilfskraft 2", 3600.0),
            ("Praktikant/in (3 Monate)", 2400.0),
            ("Externe Beratung", 8000.0),
        ];
        for (pos, (desc, cost)) in personnel.iter().enumerate() {
            values.set_position_row(
                1,
                (pos + 1) as u16,
                *desc,
                *cost,
                *cost / 2.0,
                *cost / 2.0,
                "",
            );
        }

        // Kategorie 2: Sachkosten (15 Positionen)
        let materials = [
            ("Laborchemikalien", 12000.0),
            ("Bueromaterial", 3000.0),
            ("Softwarelizenzen", 8000.0),
            ("Wartung Laborgeraete", 5000.0),
            ("Verbrauchsmaterialien", 4000.0),
            ("Kalibrierstandards", 3500.0),
            ("Abfallentsorgung", 2000.0),
            ("Energiekosten", 1500.0),
            ("Messinstrumente", 6000.0),
            ("Ersatzteile", 2500.0),
            ("Cloud-Speicher", 1200.0),
            ("Druckerpatronen", 800.0),
            ("Schulungsmaterial", 1500.0),
            ("Versicherungen", 2000.0),
            ("Laborzubehoer", 1000.0),
        ];
        for (pos, (desc, cost)) in materials.iter().enumerate() {
            values.set_position_row(
                2,
                (pos + 1) as u16,
                *desc,
                *cost,
                *cost / 2.0,
                *cost / 2.0,
                "",
            );
        }

        // Kategorie 3: Reisekosten (8 Positionen)
        let travels = [
            ("Konferenz Berlin", 3500.0),
            ("Fachtagung Muenchen", 2800.0),
            ("Kooperation Heidelberg", 1800.0),
            ("Workshop Koeln", 2200.0),
            ("Gastwissenschaftler", 1500.0),
            ("Partnerinstitut", 2000.0),
            ("Konferenzbeitraege", 1800.0),
            ("Lokale Dienstreisen", 1200.0),
        ];
        for (pos, (desc, cost)) in travels.iter().enumerate() {
            values.set_position_row(
                3,
                (pos + 1) as u16,
                *desc,
                *cost,
                *cost / 2.0,
                *cost / 2.0,
                "",
            );
        }

        // Kategorie 4: Investitionen (5 Positionen)
        let investments = [
            ("Spektrometer", 25000.0),
            ("Zentrifuge", 8000.0),
            ("Analysewaage", 4500.0),
            ("Thermocycler", 6000.0),
            ("Server", 5000.0),
        ];
        for (pos, (desc, cost)) in investments.iter().enumerate() {
            values.set_position_row(
                4,
                (pos + 1) as u16,
                *desc,
                *cost,
                *cost / 2.0,
                *cost / 2.0,
                "",
            );
        }

        // Kategorie 5: Sonstige (3 Positionen)
        let misc = [
            ("Publikationskosten", 3000.0),
            ("Fachverbaende", 1500.0),
            ("Schulungen", 800.0),
        ];
        for (pos, (desc, cost)) in misc.iter().enumerate() {
            values.set_position_row(
                5,
                (pos + 1) as u16,
                *desc,
                *cost,
                *cost / 2.0,
                *cost / 2.0,
                "",
            );
        }

        // Kategorien 6-8: Header-Input Mode
        values.set_header_input(6, 8000.0, 4000.0, 4000.0, "Raumkosten");
        values.set_header_input(7, 5000.0, 2500.0, 2500.0, "");
        values.set_header_input(8, 3000.0, 1500.0, 1500.0, "Reserve");

        // Footer
        values
            .set_footer_bank(15000.0)
            .set_footer_kasse(2500.0)
            .set_footer_sonstiges(500.0);

        let body_config = BodyConfig::new()
            .with_positions(1, 10)
            .with_positions(2, 15)
            .with_positions(3, 8)
            .with_positions(4, 5)
            .with_positions(5, 3)
            .with_positions(6, 0)
            .with_positions(7, 0)
            .with_positions(8, 0);

        let options = ReportOptions::with_default_protection().with_hidden_columns_qv();

        write_report_with_options(ws, &styles, "_de", &values, &body_config, &options)
            .expect("Failed to write report");

        workbook
            .save("tests/output/test_medium_project.xlsx")
            .expect("Failed to save");

        println!("Mittleres Projekt: tests/output/test_medium_project.xlsx");
    }

    // ========================================================================
    // Test 2: Grosses Projekt - Maximale Positionen
    // ========================================================================
    #[test]
    fn test_large_project() {
        let sheet_name = "Finanzbericht";
        let mut workbook = create_workbook_with_sheet(sheet_name);
        setup_worksheet(&mut workbook, sheet_name);

        let ws = workbook.worksheet_from_name(sheet_name).unwrap();
        let styles = ReportStyles::new();

        let mut values = ReportValues::new()
            .with_language("deutsch")
            .with_currency("EUR")
            .with_project_number("KMW-2024-LARGE")
            .with_project_title("Multi-Standort Forschungsprojekt")
            .with_project_start("01.01.2024")
            .with_project_end("31.12.2025")
            .with_report_start("01.01.2024")
            .with_report_end("30.06.2024");

        // Tabellen-Daten
        let budgets = [150000.0, 100000.0, 50000.0, 80000.0, 20000.0];
        for i in 0..5u8 {
            values
                .set(ApiKey::ApprovedBudget(i), budgets[i as usize])
                .set(
                    ApiKey::IncomeReportPeriod(i),
                    budgets[i as usize] / 2.0 * 1.05,
                )
                .set(ApiKey::IncomeTotal(i), budgets[i as usize] / 2.0)
                .set(ApiKey::IncomeReason(i), format!("Foerderline {}", i + 1));
        }

        // Right Panel
        for i in 0..18u8 {
            let amount = 2000.0 + i as f64 * 300.0;
            let date = format!("2024-{:02}-{:02}", ((i / 6) % 6) + 1, ((i % 6) * 5) + 1);
            values
                .set(ApiKey::LeftDate(i), date.clone())
                .set(ApiKey::LeftAmountEuro(i), amount)
                .set(ApiKey::LeftAmountLocal(i), amount)
                .set(ApiKey::RightDate(i), date)
                .set(ApiKey::RightAmountEuro(i), amount * 1.15)
                .set(ApiKey::RightAmountLocal(i), amount * 1.15);
        }

        // Dynamische Positionen generieren
        for pos in 1..=20u16 {
            let cost = 60000.0 - (pos as f64 - 1.0) * 2000.0;
            values.set_position_row(
                1,
                pos,
                format!("Personal {}", pos),
                cost,
                cost / 2.0,
                cost / 2.0,
                "",
            );
        }
        for pos in 1..=30u16 {
            let cost = 10000.0 - (pos as f64 - 1.0) * 300.0;
            values.set_position_row(
                2,
                pos,
                format!("Material {}", pos),
                cost,
                cost / 2.0,
                cost / 2.0,
                "",
            );
        }
        for pos in 1..=15u16 {
            let cost = 5000.0 - (pos as f64 - 1.0) * 200.0;
            values.set_position_row(
                3,
                pos,
                format!("Reise {}", pos),
                cost,
                cost / 2.0,
                cost / 2.0,
                "",
            );
        }
        for pos in 1..=10u16 {
            let cost = 20000.0 - (pos as f64 - 1.0) * 1500.0;
            values.set_position_row(
                4,
                pos,
                format!("Investition {}", pos),
                cost,
                cost / 2.0,
                cost / 2.0,
                "",
            );
        }
        for pos in 1..=5u16 {
            let cost = 8000.0 - (pos as f64 - 1.0) * 1000.0;
            values.set_position_row(
                5,
                pos,
                format!("Sonstiges {}", pos),
                cost,
                cost / 2.0,
                cost / 2.0,
                "",
            );
        }

        values.set_header_input(6, 12000.0, 6000.0, 6000.0, "Infrastruktur");
        values.set_header_input(7, 8000.0, 4000.0, 4000.0, "");
        values.set_header_input(8, 5000.0, 2500.0, 2500.0, "");

        values
            .set_footer_bank(85000.0)
            .set_footer_kasse(12000.0)
            .set_footer_sonstiges(3000.0);

        let body_config = BodyConfig::new()
            .with_positions(1, 20)
            .with_positions(2, 30)
            .with_positions(3, 15)
            .with_positions(4, 10)
            .with_positions(5, 5)
            .with_positions(6, 0)
            .with_positions(7, 0)
            .with_positions(8, 0);

        let options = ReportOptions::with_default_protection().with_hidden_columns_qv();

        write_report_with_options(ws, &styles, "_de", &values, &body_config, &options)
            .expect("Failed to write report");

        workbook
            .save("tests/output/test_large_project.xlsx")
            .expect("Failed to save");

        println!("Grosses Projekt: tests/output/test_large_project.xlsx");
    }

    // ========================================================================
    // Test 3: Minimales Projekt
    // ========================================================================
    #[test]
    fn test_minimal_project() {
        let sheet_name = "Finanzbericht";
        let mut workbook = create_workbook_with_sheet(sheet_name);
        setup_worksheet(&mut workbook, sheet_name);

        let ws = workbook.worksheet_from_name(sheet_name).unwrap();
        let styles = ReportStyles::new();

        let mut values = ReportValues::new()
            .with_language("deutsch")
            .with_currency("EUR")
            .with_project_number("KMW-2024-MINI")
            .with_project_title("Kleinprojekt")
            .with_project_start("01.06.2024")
            .with_project_end("31.08.2024")
            .with_report_start("01.06.2024")
            .with_report_end("30.06.2024");

        for i in 0..5u8 {
            values
                .set(ApiKey::ApprovedBudget(i), 10000.0 + i as f64 * 2000.0)
                .set(ApiKey::IncomeReportPeriod(i), 5000.0 + i as f64 * 1000.0)
                .set(ApiKey::IncomeTotal(i), 5000.0 + i as f64 * 1000.0)
                .set(ApiKey::IncomeReason(i), format!("Foerderung {}", i + 1));
        }

        for i in 0..18u8 {
            let amount = 800.0 + i as f64 * 150.0;
            let date = format!("2024-06-{:02}", (i % 30) + 1);
            values
                .set(ApiKey::LeftDate(i), date.clone())
                .set(ApiKey::LeftAmountEuro(i), amount)
                .set(ApiKey::LeftAmountLocal(i), amount)
                .set(ApiKey::RightDate(i), date)
                .set(ApiKey::RightAmountEuro(i), amount)
                .set(ApiKey::RightAmountLocal(i), amount);
        }

        // Wenige Positionen
        values.set_position_row(1, 1, "Projektmanagement", 15000.0, 7500.0, 7500.0, "");
        values.set_position_row(
            1,
            2,
            "Wissenschaftliches Personal",
            12000.0,
            6000.0,
            6000.0,
            "",
        );
        values.set_position_row(2, 1, "Materialien", 8000.0, 4000.0, 4000.0, "");
        values.set_position_row(2, 2, "Software", 4000.0, 2000.0, 2000.0, "");
        values.set_position_row(3, 1, "Konferenzreisen", 4000.0, 2000.0, 2000.0, "");

        // Alle anderen Header-Input
        values.set_header_input(4, 5000.0, 2500.0, 2500.0, "");
        values.set_header_input(5, 2000.0, 1000.0, 1000.0, "");
        values.set_header_input(6, 3000.0, 1500.0, 1500.0, "");
        values.set_header_input(7, 2000.0, 1000.0, 1000.0, "");
        values.set_header_input(8, 1000.0, 500.0, 500.0, "");

        let body_config = BodyConfig::new()
            .with_positions(1, 2)
            .with_positions(2, 2)
            .with_positions(3, 1)
            .with_positions(4, 0)
            .with_positions(5, 0)
            .with_positions(6, 0)
            .with_positions(7, 0)
            .with_positions(8, 0);

        let options = ReportOptions::with_default_protection().with_hidden_columns_qv();

        write_report_with_options(ws, &styles, "_de", &values, &body_config, &options)
            .expect("Failed to write report");

        workbook
            .save("tests/output/test_minimal_project.xlsx")
            .expect("Failed to save");

        println!("Minimales Projekt: tests/output/test_minimal_project.xlsx");
    }

    // ========================================================================
    // Test 4: Englische Version
    // ========================================================================
    #[test]
    fn test_english_version() {
        let sheet_name = "Financial Report";
        let mut workbook = create_workbook_with_sheet(sheet_name);
        setup_worksheet(&mut workbook, sheet_name);

        let ws = workbook.worksheet_from_name(sheet_name).unwrap();
        let styles = ReportStyles::new();

        let mut values = ReportValues::new()
            .with_language("englisch")
            .with_currency("USD")
            .with_project_number("KMW-2024-EN")
            .with_project_title("International Research Project")
            .with_project_start("01/01/2024")
            .with_project_end("12/31/2024")
            .with_report_start("01/01/2024")
            .with_report_end("06/30/2024");

        for i in 0..5u8 {
            values
                .set(ApiKey::ApprovedBudget(i), 40000.0 + i as f64 * 10000.0)
                .set(ApiKey::IncomeReportPeriod(i), 20000.0 + i as f64 * 5000.0)
                .set(ApiKey::IncomeTotal(i), 20000.0 + i as f64 * 5000.0)
                .set(ApiKey::IncomeReason(i), format!("Funding source {}", i + 1));
        }

        for i in 0..18u8 {
            let amount = 1200.0 + i as f64 * 180.0;
            let date = format!("{:02}/15/2024", (i % 12) + 1);
            values
                .set(ApiKey::LeftDate(i), date.clone())
                .set(ApiKey::LeftAmountEuro(i), amount)
                .set(ApiKey::LeftAmountLocal(i), amount * 1.08)
                .set(ApiKey::RightDate(i), date)
                .set(ApiKey::RightAmountEuro(i), amount * 1.1)
                .set(ApiKey::RightAmountLocal(i), amount * 1.18);
        }

        // English descriptions
        values.set_position_row(
            1,
            1,
            "Project Manager (100%)",
            50000.0,
            25000.0,
            25000.0,
            "",
        );
        values.set_position_row(
            1,
            2,
            "Senior Researcher (100%)",
            45000.0,
            22500.0,
            22500.0,
            "",
        );
        values.set_position_row(
            1,
            3,
            "Research Assistant (80%)",
            28000.0,
            14000.0,
            14000.0,
            "",
        );
        values.set_position_row(2, 1, "Laboratory Supplies", 10000.0, 5000.0, 5000.0, "");
        values.set_position_row(2, 2, "Software Licenses", 6000.0, 3000.0, 3000.0, "");
        values.set_position_row(3, 1, "Conference Travel", 4000.0, 2000.0, 2000.0, "");
        values.set_position_row(3, 2, "Partner Visits", 3000.0, 1500.0, 1500.0, "");
        values.set_position_row(4, 1, "Lab Equipment", 15000.0, 7500.0, 7500.0, "");

        values.set_header_input(5, 5000.0, 2500.0, 2500.0, "Miscellaneous");
        values.set_header_input(6, 6000.0, 3000.0, 3000.0, "Facilities");
        values.set_header_input(7, 4000.0, 2000.0, 2000.0, "");
        values.set_header_input(8, 2000.0, 1000.0, 1000.0, "Contingency");

        values
            .set_footer_bank(12000.0)
            .set_footer_kasse(1500.0)
            .set_footer_sonstiges(300.0);

        let body_config = BodyConfig::new()
            .with_positions(1, 3)
            .with_positions(2, 2)
            .with_positions(3, 2)
            .with_positions(4, 1)
            .with_positions(5, 0)
            .with_positions(6, 0)
            .with_positions(7, 0)
            .with_positions(8, 0);

        let options = ReportOptions::with_default_protection().with_hidden_columns_qv();

        write_report_with_options(ws, &styles, "_en", &values, &body_config, &options)
            .expect("Failed to write report");

        workbook
            .save("tests/output/test_english_version.xlsx")
            .expect("Failed to save");

        println!("English Version: tests/output/test_english_version.xlsx");
    }

    // ========================================================================
    // Test 5: Kategorie 6-8 mit Positionen (Ausnahmefall)
    // ========================================================================
    #[test]
    fn test_categories_with_positions() {
        let sheet_name = "Finanzbericht";
        let mut workbook = create_workbook_with_sheet(sheet_name);
        setup_worksheet(&mut workbook, sheet_name);

        let ws = workbook.worksheet_from_name(sheet_name).unwrap();
        let styles = ReportStyles::new();

        let mut values = ReportValues::new()
            .with_language("deutsch")
            .with_currency("EUR")
            .with_project_number("KMW-2024-SPECIAL")
            .with_project_title("Projekt mit speziellen Kategorien")
            .with_project_start("01.01.2024")
            .with_project_end("31.12.2024")
            .with_report_start("01.01.2024")
            .with_report_end("30.06.2024");

        for i in 0..5u8 {
            values
                .set(ApiKey::ApprovedBudget(i), 30000.0 + i as f64 * 5000.0)
                .set(ApiKey::IncomeReportPeriod(i), 15000.0 + i as f64 * 2500.0)
                .set(ApiKey::IncomeTotal(i), 15000.0 + i as f64 * 2500.0)
                .set(ApiKey::IncomeReason(i), format!("Quelle {}", i + 1));
        }

        for i in 0..18u8 {
            let amount = 1500.0 + i as f64 * 200.0;
            let date = format!("2024-{:02}-10", ((i / 3) % 12) + 1);
            values
                .set(ApiKey::LeftDate(i), date.clone())
                .set(ApiKey::LeftAmountEuro(i), amount)
                .set(ApiKey::LeftAmountLocal(i), amount)
                .set(ApiKey::RightDate(i), date)
                .set(ApiKey::RightAmountEuro(i), amount)
                .set(ApiKey::RightAmountLocal(i), amount);
        }

        // Standard-Kategorien
        for pos in 1..=8u16 {
            values.set_position_row(
                1,
                pos,
                format!("Personal {}", pos),
                25000.0,
                12500.0,
                12500.0,
                "",
            );
        }
        for pos in 1..=10u16 {
            values.set_position_row(
                2,
                pos,
                format!("Sachkosten {}", pos),
                5000.0,
                2500.0,
                2500.0,
                "",
            );
        }
        for pos in 1..=5u16 {
            values.set_position_row(
                3,
                pos,
                format!("Reisen {}", pos),
                3000.0,
                1500.0,
                1500.0,
                "",
            );
        }
        for pos in 1..=3u16 {
            values.set_position_row(
                4,
                pos,
                format!("Investition {}", pos),
                10000.0,
                5000.0,
                5000.0,
                "",
            );
        }
        for pos in 1..=2u16 {
            values.set_position_row(
                5,
                pos,
                format!("Sonstiges {}", pos),
                2500.0,
                1250.0,
                1250.0,
                "",
            );
        }

        // Kategorie 6: Mit Positionen (Ausnahme)
        values.set_position_row(6, 1, "Spezialausstattung A", 12000.0, 6000.0, 6000.0, "");
        values.set_position_row(6, 2, "Spezialausstattung B", 8000.0, 4000.0, 4000.0, "");

        // Kategorie 7: Header-Input
        values.set_header_input(7, 10000.0, 5000.0, 5000.0, "Verwaltung");

        // Kategorie 8: Mit Positionen (Reserve-Detail)
        values.set_position_row(
            8,
            1,
            "Notfallreserve",
            3000.0,
            1500.0,
            1500.0,
            "Unvorhergesehenes",
        );
        values.set_position_row(8, 2, "Preisanpassung", 2000.0, 1000.0, 1000.0, "Inflation");
        values.set_position_row(8, 3, "Zusatzpersonal", 2000.0, 1000.0, 1000.0, "Vertretung");

        values
            .set_footer_bank(25000.0)
            .set_footer_kasse(3000.0)
            .set_footer_sonstiges(500.0);

        let body_config = BodyConfig::new()
            .with_positions(1, 8)
            .with_positions(2, 10)
            .with_positions(3, 5)
            .with_positions(4, 3)
            .with_positions(5, 2)
            .with_positions(6, 2) // Ausnahme: 2 Positionen
            .with_positions(7, 0) // Header-Input
            .with_positions(8, 3); // Ausnahme: 3 Positionen

        let options = ReportOptions::with_default_protection().with_hidden_columns_qv();

        write_report_with_options(ws, &styles, "_de", &values, &body_config, &options)
            .expect("Failed to write report");

        workbook
            .save("tests/output/test_special_categories.xlsx")
            .expect("Failed to save");

        println!("Spezielle Kategorien: tests/output/test_special_categories.xlsx");
    }
}
