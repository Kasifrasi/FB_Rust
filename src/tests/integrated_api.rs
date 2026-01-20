//! Integrierte Tests: Flexible Body-API + Right Panel + Header
//!
//! Testet realistische Szenarien mit:
//! - Dynamischen Kostenpositionen (10-30 pro Kategorie)
//! - Kategorien 6-8: Standardmäßig Header-Input Mode (position=0)
//! - Right Panel Daten (L, M, N, S, T, U Spalten)
//! - Statischen Header und Tabellen-Daten

#[cfg(test)]
mod integrated_api_tests {
    use crate::report::api::ApiKey;
    use crate::report::formats::ReportStyles;
    use crate::report::layout::setup_sheet;
    use crate::report::values::ReportValues;
    use crate::report::writer::write_report_with_body;
    use crate::report::BodyConfig;
    use crate::lang::builder::build_sheet as build_trans_sheet;
    use rust_xlsxwriter::{Format, Workbook};

    // ========================================================================
    // Test 1: Realistische Projektbudgetierung - Mittelgroßes Projekt
    // ========================================================================

    /// Test 1: Mittleres Forschungsprojekt mit realistischen Kosten
    ///
    /// Konfiguration:
    /// - Kategorie 1 (Personal): 10 Positionen
    /// - Kategorie 2 (Sachkosten): 15 Positionen
    /// - Kategorie 3 (Reisekosten): 8 Positionen
    /// - Kategorie 4 (Investitionen): 5 Positionen
    /// - Kategorie 5 (Sonstige): 3 Positionen
    /// - Kategorien 6-8: Header-Input Mode (position=0)
    ///
    /// Realistische Kostenverteilung mit vollständigen Daten
    #[test]
    fn test_realistic_medium_project() {
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

        // === STATISCHE WERTE ===
        let mut values = ReportValues::new()
            .with_language("deutsch")
            .with_currency("EUR")
            .with_project_number("KMW-2024-001")
            .with_project_title("Mittelpunktanalyse von Forschungsmethoden")
            .with_project_start("01.01.2024")
            .with_project_end("31.12.2024")
            .with_report_start("01.01.2024")
            .with_report_end("30.06.2024");

        // Tabellen-Daten (D15-F19, H15-H19)
        let approved_budgets = [50000.0, 30000.0, 15000.0, 8000.0, 5000.0];
        let income_periods = [25000.0, 15000.0, 7500.0, 4000.0, 2500.0];
        for i in 0..5u8 {
            values.set(ApiKey::ApprovedBudget(i), approved_budgets[i as usize]);
            values.set(
                ApiKey::IncomeReportPeriod(i),
                income_periods[i as usize] * 1.1,
            );
            values.set(ApiKey::IncomeTotal(i), income_periods[i as usize]);
            values.set(
                ApiKey::IncomeReason(i),
                format!("Finanzierungslinie {}", i + 1),
            );
        }

        // Right Panel Daten (L14-N31 und S14-U31) - 18 Zeilen
        for i in 0..18u8 {
            let base_amount = 1000.0 + i as f64 * 200.0;
            let date = format!("2024-{:02}-{:02}", ((i / 6) % 6) + 1, ((i % 6) * 5) + 1);

            values
                .set(ApiKey::LeftDate(i), date.clone())
                .set(ApiKey::LeftAmountEuro(i), base_amount)
                .set(ApiKey::LeftAmountLocal(i), base_amount)
                .set(ApiKey::RightDate(i), date)
                .set(ApiKey::RightAmountEuro(i), base_amount * 1.2)
                .set(ApiKey::RightAmountLocal(i), base_amount * 1.2);
        }

        // === DYNAMISCHE WERTE: Kostenpositionen ===

        // Kategorie 1: 10 Personalkosten
        let personnel = [
            ("Projektleiter (100%)", 60000.0, 30000.0),
            ("Senior Wissenschaftler (100%)", 55000.0, 27500.0),
            ("Wissenschaftlicher Mitarbeiter 1 (80%)", 32000.0, 16000.0),
            ("Wissenschaftlicher Mitarbeiter 2 (80%)", 32000.0, 16000.0),
            ("Techniker/in (50%)", 18000.0, 9000.0),
            ("Sekretariat/Administration (30%)", 10000.0, 5000.0),
            ("Studentische Hilfskraft 1 (20h/Woche)", 3600.0, 1800.0),
            ("Studentische Hilfskraft 2 (20h/Woche)", 3600.0, 1800.0),
            ("Praktikant/in (3 Monate)", 2400.0, 1200.0),
            ("Externe Beratung (Teilzeit)", 8000.0, 4000.0),
        ];

        for (pos, (desc, approved, income)) in personnel.iter().enumerate() {
            let pos = (pos + 1) as u16;
            values.set_position_row(1, pos, *desc, *approved, *income, *income, "");
        }

        // Kategorie 2: 15 Sachkosten
        let materials = [
            ("Laborchemikalien und Reagenzien", 12000.0, 6000.0),
            ("Büromaterial und Druckkosten", 3000.0, 1500.0),
            ("Softwarelizenzen (Statistik, CAD)", 8000.0, 4000.0),
            ("Wartung Laborgeräte", 5000.0, 2500.0),
            ("Verbrauchsmaterialien (Pipetten, Behälter)", 4000.0, 2000.0),
            ("Kalibrierstandards und Referenzmaterialien", 3500.0, 1750.0),
            ("Abfallwirtschaft und Entsorgung", 2000.0, 1000.0),
            ("Energiekosten Labor (Zusatz)", 1500.0, 750.0),
            ("Beschaffung Messinstrumente", 6000.0, 3000.0),
            ("Ersatzteile Geräte", 2500.0, 1250.0),
            ("Datenträgerverwaltung und Cloud", 1200.0, 600.0),
            ("Druckerpatronen und Toner", 800.0, 400.0),
            ("Lehr- und Schulungsmaterialien", 1500.0, 750.0),
            ("Versicherungen Laborgeräte", 2000.0, 1000.0),
            ("Sonstiges Laborzubehör", 1000.0, 500.0),
        ];

        for (pos, (desc, approved, income)) in materials.iter().enumerate() {
            let pos = (pos + 1) as u16;
            values.set_position_row(2, pos, *desc, *approved, *income, *income, "");
        }

        // Kategorie 3: 8 Reisekosten
        let travels = [
            ("Konferenz 'Methods in Research' (Berlin)", 3500.0, 1750.0),
            ("Fachtagung Laborstandards (München)", 2800.0, 1400.0),
            ("Kooperationstreffen Universität Heidelberg", 1800.0, 900.0),
            ("Workshop Datenanalyse (Köln)", 2200.0, 1100.0),
            ("Gastwissenschaftler-Besuch", 1500.0, 750.0),
            ("Forschungsfahrt zu Partnerinstitut (4x)", 2000.0, 1000.0),
            ("Konferenzbeiträge anmelden und präsentieren", 1800.0, 900.0),
            ("Lokale Dienstreisen (Taxi, Bahn)", 1200.0, 600.0),
        ];

        for (pos, (desc, approved, income)) in travels.iter().enumerate() {
            let pos = (pos + 1) as u16;
            values.set_position_row(3, pos, *desc, *approved, *income, *income, "");
        }

        // Kategorie 4: 5 Investitionen
        let investments = [
            ("Hochleistungs-Spektrometer", 25000.0, 12500.0),
            ("Zentrifuge (6000 U/min)", 8000.0, 4000.0),
            ("Analysewaage Präzision ±0.0001g", 4500.0, 2250.0),
            ("Thermocycler für PCR", 6000.0, 3000.0),
            ("Laborcomputer und Server", 5000.0, 2500.0),
        ];

        for (pos, (desc, approved, income)) in investments.iter().enumerate() {
            let pos = (pos + 1) as u16;
            values.set_position_row(4, pos, *desc, *approved, *income, *income, "");
        }

        // Kategorie 5: 3 Sonstige
        let misc = [
            ("Publikationskosten Open Access", 3000.0, 1500.0),
            ("Mitgliedschaften Fachverbände", 1500.0, 750.0),
            ("Sicherheitsunterweisungen und Schulungen", 800.0, 400.0),
        ];

        for (pos, (desc, approved, income)) in misc.iter().enumerate() {
            let pos = (pos + 1) as u16;
            values.set_position_row(5, pos, *desc, *approved, *income, *income, "");
        }

        // Kategorien 6-8: Header-Input Mode (nur position=0)
        values.set_header_input(6, 8000.0, 4000.0, 4000.0, "Raumsummen und Nebenkosten");
        values.set_header_input(7, 5000.0, 2500.0, 2500.0, "");
        values.set_header_input(8, 3000.0, 1500.0, 1500.0, "Reserve für Unvorhergesehenes");

        // Footer-Werte (Saldenabstimmung)
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
            .with_positions(6, 0) // Header-Input
            .with_positions(7, 0) // Header-Input
            .with_positions(8, 0); // Header-Input

        write_report_with_body(ws, &styles, "_de", &values, &body_config)
            .expect("Failed to write report");

        ws.protect();

        let path = "tests/output/test_realistic_medium_project.xlsx";
        workbook.save(path).expect("Failed to save workbook");

        println!("✓ Test 1: Realistisches mittleres Projekt");
        println!("   Datei: {}", path);
        println!("   - Kategorie 1: 10 Personalkosten");
        println!("   - Kategorie 2: 15 Sachkosten");
        println!("   - Kategorie 3: 8 Reisekosten");
        println!("   - Kategorie 4: 5 Investitionen");
        println!("   - Kategorie 5: 3 Sonstige");
        println!("   - Kategorien 6-8: Header-Input Mode");
    }

    // ========================================================================
    // Test 2: Großes Projekt mit maximalen Kostenpositionen
    // ========================================================================

    /// Test 2: Großprojekt mit detaillierten Kostenpositionen
    ///
    /// Konfiguration:
    /// - Kategorie 1 (Personal): 20 Positionen
    /// - Kategorie 2 (Sachkosten): 30 Positionen
    /// - Kategorie 3 (Reisekosten): 15 Positionen
    /// - Kategorie 4 (Investitionen): 10 Positionen
    /// - Kategorie 5 (Sonstige): 5 Positionen
    /// - Kategorien 6-8: Header-Input Mode
    #[test]
    fn test_large_project_maximum_positions() {
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

        // === STATISCHE WERTE ===
        let mut values = ReportValues::new()
            .with_language("deutsch")
            .with_currency("EUR")
            .with_project_number("KMW-2024-LARGE")
            .with_project_title("Umfassendes Multi-Standort Forschungsprojekt")
            .with_project_start("01.01.2024")
            .with_project_end("31.12.2025")
            .with_report_start("01.01.2024")
            .with_report_end("30.06.2024");

        // Tabellen-Daten
        let approved_budgets = [150000.0, 100000.0, 50000.0, 80000.0, 20000.0];
        let income_periods = [75000.0, 50000.0, 25000.0, 40000.0, 10000.0];
        for i in 0..5u8 {
            values.set(ApiKey::ApprovedBudget(i), approved_budgets[i as usize]);
            values.set(
                ApiKey::IncomeReportPeriod(i),
                income_periods[i as usize] * 1.05,
            );
            values.set(ApiKey::IncomeTotal(i), income_periods[i as usize]);
            values.set(ApiKey::IncomeReason(i), format!("Förderline {}", i + 1));
        }

        // Right Panel Daten
        for i in 0..18u8 {
            let base_amount = 2000.0 + i as f64 * 300.0;
            let date = format!("2024-{:02}-{:02}", ((i / 6) % 6) + 1, ((i % 6) * 5) + 1);

            values
                .set(ApiKey::LeftDate(i), date.clone())
                .set(ApiKey::LeftAmountEuro(i), base_amount)
                .set(ApiKey::LeftAmountLocal(i), base_amount)
                .set(ApiKey::RightDate(i), date)
                .set(ApiKey::RightAmountEuro(i), base_amount * 1.15)
                .set(ApiKey::RightAmountLocal(i), base_amount * 1.15);
        }

        // === DYNAMISCHE WERTE: 20 Personalkosten ===
        for pos in 1..=20u16 {
            let cost = 60000.0 - (pos as f64 - 1.0) * 2000.0;
            let income = cost / 2.0;
            values.set_position_row(
                1,
                pos,
                format!("Person {}: Rolle und Aufgaben", pos),
                cost,
                income,
                income,
                "",
            );
        }

        // === DYNAMISCHE WERTE: 30 Sachkosten ===
        for pos in 1..=30u16 {
            let cost = 10000.0 - (pos as f64 - 1.0) * 300.0;
            let income = cost / 2.0;
            values.set_position_row(
                2,
                pos,
                format!("Material/Service {}", pos),
                cost,
                income,
                income,
                "",
            );
        }

        // === DYNAMISCHE WERTE: 15 Reisekosten ===
        for pos in 1..=15u16 {
            let cost = 5000.0 - (pos as f64 - 1.0) * 200.0;
            let income = cost / 2.0;
            values.set_position_row(3, pos, format!("Reise {}", pos), cost, income, income, "");
        }

        // === DYNAMISCHE WERTE: 10 Investitionen ===
        for pos in 1..=10u16 {
            let cost = 20000.0 - (pos as f64 - 1.0) * 1500.0;
            let income = cost / 2.0;
            values.set_position_row(
                4,
                pos,
                format!("Investition {}", pos),
                cost,
                income,
                income,
                "",
            );
        }

        // === DYNAMISCHE WERTE: 5 Sonstige ===
        for pos in 1..=5u16 {
            let cost = 8000.0 - (pos as f64 - 1.0) * 1000.0;
            let income = cost / 2.0;
            values.set_position_row(
                5,
                pos,
                format!("Sonstiges {}", pos),
                cost,
                income,
                income,
                "",
            );
        }

        // Kategorien 6-8: Header-Input Mode
        values.set_header_input(6, 12000.0, 6000.0, 6000.0, "Infrastruktur");
        values.set_header_input(7, 8000.0, 4000.0, 4000.0, "");
        values.set_header_input(8, 5000.0, 2500.0, 2500.0, "");

        // Footer-Werte (Saldenabstimmung) - Großprojekt
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
            .with_positions(6, 0) // Header-Input
            .with_positions(7, 0) // Header-Input
            .with_positions(8, 0); // Header-Input

        write_report_with_body(ws, &styles, "_de", &values, &body_config)
            .expect("Failed to write report");

        ws.protect();

        let path = "tests/output/test_large_project_maximum.xlsx";
        workbook.save(path).expect("Failed to save workbook");

        println!("✓ Test 2: Großprojekt mit maximalen Positionen");
        println!("   Datei: {}", path);
        println!("   - Kategorie 1: 20 Positionen");
        println!("   - Kategorie 2: 30 Positionen");
        println!("   - Kategorie 3: 15 Positionen");
        println!("   - Kategorie 4: 10 Positionen");
        println!("   - Kategorie 5: 5 Positionen");
        println!("   - Kategorien 6-8: Header-Input Mode");
    }

    // ========================================================================
    // Test 3: Minimales Projekt - nur Header-Input für 6-8
    // ========================================================================

    /// Test 3: Minimales Projekt
    ///
    /// Konfiguration:
    /// - Kategorie 1: 5 Positionen
    /// - Kategorie 2: 4 Positionen
    /// - Kategorie 3: 3 Positionen
    /// - Kategorien 4-8: Header-Input Mode
    #[test]
    fn test_minimal_project() {
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

        // === STATISCHE WERTE ===
        let mut values = ReportValues::new()
            .with_language("deutsch")
            .with_currency("EUR")
            .with_project_number("KMW-2024-MINI")
            .with_project_title("Kleinprojekt - Kostenpilot")
            .with_project_start("01.06.2024")
            .with_project_end("31.08.2024")
            .with_report_start("01.06.2024")
            .with_report_end("30.06.2024");

        // Tabellen-Daten
        for i in 0..5u8 {
            values.set(ApiKey::ApprovedBudget(i), 10000.0 + i as f64 * 2000.0);
            values.set(ApiKey::IncomeReportPeriod(i), 5000.0 + i as f64 * 1000.0);
            values.set(ApiKey::IncomeTotal(i), 5000.0 + i as f64 * 1000.0);
            values.set(ApiKey::IncomeReason(i), format!("Förderung {}", i + 1));
        }

        // Right Panel Daten
        for i in 0..18u8 {
            let base_amount = 800.0 + i as f64 * 150.0;
            let date = format!("2024-06-{:02}", (i % 30) + 1);

            values
                .set(ApiKey::LeftDate(i), date.clone())
                .set(ApiKey::LeftAmountEuro(i), base_amount)
                .set(ApiKey::LeftAmountLocal(i), base_amount)
                .set(ApiKey::RightDate(i), date)
                .set(ApiKey::RightAmountEuro(i), base_amount)
                .set(ApiKey::RightAmountLocal(i), base_amount);
        }

        // === DYNAMISCHE WERTE ===

        // Kategorie 1: 5 Positionen
        let costs1 = [
            ("Personal Projektmanagement", 15000.0),
            ("Wissenschaftliches Personal", 12000.0),
            ("Administrative Unterstützung", 5000.0),
            ("Studentische Hilfskraft", 2000.0),
            ("Externe Beratung", 3000.0),
        ];

        for (pos, (desc, cost)) in costs1.iter().enumerate() {
            let pos = (pos + 1) as u16;
            values.set_position_row(1, pos, *desc, *cost, *cost / 2.0, *cost / 2.0, "");
        }

        // Kategorie 2: 4 Positionen
        let costs2 = [
            ("Materialien und Verbrauchsstoffe", 8000.0),
            ("Software und IT", 4000.0),
            ("Ausrüstungsverwaltung", 3000.0),
            ("Dokumentation und Reporting", 2000.0),
        ];

        for (pos, (desc, cost)) in costs2.iter().enumerate() {
            let pos = (pos + 1) as u16;
            values.set_position_row(2, pos, *desc, *cost, *cost / 2.0, *cost / 2.0, "");
        }

        // Kategorie 3: 3 Positionen
        let costs3 = [
            ("Konferenzreisen", 4000.0),
            ("Lokale Dienstreisen", 1500.0),
            ("Partnerbesuche", 1500.0),
        ];

        for (pos, (desc, cost)) in costs3.iter().enumerate() {
            let pos = (pos + 1) as u16;
            values.set_position_row(3, pos, *desc, *cost, *cost / 2.0, *cost / 2.0, "");
        }

        // Kategorien 4-8: Header-Input Mode
        values.set_header_input(4, 5000.0, 2500.0, 2500.0, "");
        values.set_header_input(5, 2000.0, 1000.0, 1000.0, "");
        values.set_header_input(6, 3000.0, 1500.0, 1500.0, "");
        values.set_header_input(7, 2000.0, 1000.0, 1000.0, "");
        values.set_header_input(8, 1000.0, 500.0, 500.0, "");

        let body_config = BodyConfig::new()
            .with_positions(1, 5)
            .with_positions(2, 4)
            .with_positions(3, 3)
            .with_positions(4, 0) // Header-Input
            .with_positions(5, 0) // Header-Input
            .with_positions(6, 0) // Header-Input
            .with_positions(7, 0) // Header-Input
            .with_positions(8, 0); // Header-Input

        write_report_with_body(ws, &styles, "_de", &values, &body_config)
            .expect("Failed to write report");

        ws.protect();

        let path = "tests/output/test_minimal_project.xlsx";
        workbook.save(path).expect("Failed to save workbook");

        println!("✓ Test 3: Minimalprojekt");
        println!("   Datei: {}", path);
        println!("   - Kategorie 1: 5 Positionen");
        println!("   - Kategorie 2: 4 Positionen");
        println!("   - Kategorie 3: 3 Positionen");
        println!("   - Kategorien 4-8: Header-Input Mode");
    }

    // ========================================================================
    // Test 4: Spezial - Kategorien 6-8 mit minimalen Positionen als Beispiel
    // ========================================================================

    /// Test 4: Kategorien 6-8 mit seltener Verwendung von Positionen
    ///
    /// Konfiguration:
    /// - Kategorien 1-5: Normale Positionen
    /// - Kategorie 6: 2 Positionen (Ausnahmefall)
    /// - Kategorie 7: Header-Input Mode
    /// - Kategorie 8: Header-Input Mode (Reserve)
    #[test]
    fn test_category_6_with_minimal_positions() {
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

        // === STATISCHE WERTE ===
        let mut values = ReportValues::new()
            .with_language("deutsch")
            .with_currency("EUR")
            .with_project_number("KMW-2024-MIXED")
            .with_project_title("Projekt mit gemischten Kategorien")
            .with_project_start("01.03.2024")
            .with_project_end("29.02.2025")
            .with_report_start("01.03.2024")
            .with_report_end("30.06.2024");

        // Tabellen-Daten
        for i in 0..5u8 {
            values.set(ApiKey::ApprovedBudget(i), 40000.0 + i as f64 * 10000.0);
            values.set(ApiKey::IncomeReportPeriod(i), 20000.0 + i as f64 * 5000.0);
            values.set(ApiKey::IncomeTotal(i), 20000.0 + i as f64 * 5000.0);
            values.set(ApiKey::IncomeReason(i), format!("Quelle {}", i + 1));
        }

        // Right Panel Daten
        for i in 0..18u8 {
            let base_amount = 1200.0 + i as f64 * 250.0;
            let date = format!("2024-{:02}-15", (i % 12) + 1);

            values
                .set(ApiKey::LeftDate(i), date.clone())
                .set(ApiKey::LeftAmountEuro(i), base_amount)
                .set(ApiKey::LeftAmountLocal(i), base_amount * 1.05)
                .set(ApiKey::RightDate(i), date)
                .set(ApiKey::RightAmountEuro(i), base_amount * 1.1)
                .set(ApiKey::RightAmountLocal(i), base_amount * 1.15);
        }

        // === DYNAMISCHE WERTE ===

        // Kategorie 1: 12 Positionen
        for pos in 1..=12u16 {
            values.set_position_row(
                1,
                pos,
                format!("Personal Kostenstelle {}", pos),
                40000.0 - (pos as f64 - 1.0) * 2000.0,
                20000.0 - (pos as f64 - 1.0) * 1000.0,
                20000.0 - (pos as f64 - 1.0) * 1000.0,
                "",
            );
        }

        // Kategorie 2: 18 Positionen
        for pos in 1..=18u16 {
            values.set_position_row(
                2,
                pos,
                format!("Material Typ {}", pos),
                8000.0 - (pos as f64 - 1.0) * 300.0,
                4000.0 - (pos as f64 - 1.0) * 150.0,
                4000.0 - (pos as f64 - 1.0) * 150.0,
                "",
            );
        }

        // Kategorie 3: 10 Positionen
        for pos in 1..=10u16 {
            values.set_position_row(
                3,
                pos,
                format!("Reiseziel {}", pos),
                6000.0 - (pos as f64 - 1.0) * 400.0,
                3000.0 - (pos as f64 - 1.0) * 200.0,
                3000.0 - (pos as f64 - 1.0) * 200.0,
                "",
            );
        }

        // Kategorie 4: 6 Positionen
        for pos in 1..=6u16 {
            values.set_position_row(
                4,
                pos,
                format!("Ausrüstung {}", pos),
                15000.0 - (pos as f64 - 1.0) * 2000.0,
                7500.0 - (pos as f64 - 1.0) * 1000.0,
                7500.0 - (pos as f64 - 1.0) * 1000.0,
                "",
            );
        }

        // Kategorie 5: 4 Positionen
        for pos in 1..=4u16 {
            values.set_position_row(
                5,
                pos,
                format!("Sonstiges {}", pos),
                5000.0 - (pos as f64 - 1.0) * 800.0,
                2500.0 - (pos as f64 - 1.0) * 400.0,
                2500.0 - (pos as f64 - 1.0) * 400.0,
                "",
            );
        }

        // Kategorie 6: 2 Positionen (seltene Ausnahme)
        values.set_position_row(6, 1, "Spezialausstattung A", 12000.0, 6000.0, 6000.0, "");
        values.set_position_row(6, 2, "Spezialausstattung B", 8000.0, 4000.0, 4000.0, "");

        // Kategorien 7-8: Header-Input Mode
        values.set_header_input(7, 10000.0, 5000.0, 5000.0, "Verwaltung und Organisation");
        values.set_header_input(8, 5000.0, 2500.0, 2500.0, "Reserve");

        let body_config = BodyConfig::new()
            .with_positions(1, 12)
            .with_positions(2, 18)
            .with_positions(3, 10)
            .with_positions(4, 6)
            .with_positions(5, 4)
            .with_positions(6, 2) // Ausnahmefall: 2 Positionen
            .with_positions(7, 0) // Header-Input
            .with_positions(8, 0); // Header-Input

        write_report_with_body(ws, &styles, "_de", &values, &body_config)
            .expect("Failed to write report");

        ws.protect();

        let path = "tests/output/test_category_6_with_positions.xlsx";
        workbook.save(path).expect("Failed to save workbook");

        println!("✓ Test 4: Kategorie 6 mit minimalen Positionen");
        println!("   Datei: {}", path);
        println!("   - Kategorien 1-5: Multiple Positionen");
        println!("   - Kategorie 6: 2 Positionen (Ausnahme)");
        println!("   - Kategorien 7-8: Header-Input Mode");
    }

    // ========================================================================
    // Test 5: Test mit Reserve als Ausnahme mit 4 Positionen
    // ========================================================================

    /// Test 5: Reserve (Kategorie 8) mit seltener Mehrposition
    ///
    /// Konfiguration:
    /// - Kategorien 1-5: Variable Positionen
    /// - Kategorien 6-7: Header-Input Mode
    /// - Kategorie 8 (Reserve): 4 Positionen (seltene Ausnahme)
    #[test]
    fn test_reserve_with_multiple_positions() {
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

        // === STATISCHE WERTE ===
        let mut values = ReportValues::new()
            .with_language("deutsch")
            .with_currency("EUR")
            .with_project_number("KMW-2024-RESERVE")
            .with_project_title("Projekt mit Reserve-Detailung")
            .with_project_start("01.01.2024")
            .with_project_end("31.12.2024")
            .with_report_start("01.01.2024")
            .with_report_end("30.06.2024");

        // Tabellen-Daten
        for i in 0..5u8 {
            values.set(ApiKey::ApprovedBudget(i), 30000.0 + i as f64 * 5000.0);
            values.set(ApiKey::IncomeReportPeriod(i), 15000.0 + i as f64 * 2500.0);
            values.set(ApiKey::IncomeTotal(i), 15000.0 + i as f64 * 2500.0);
            values.set(
                ApiKey::IncomeReason(i),
                format!("Finanzierungsquelle {}", i + 1),
            );
        }

        // Right Panel Daten
        for i in 0..18u8 {
            let base_amount = 1500.0 + i as f64 * 200.0;
            let date = format!("2024-{:02}-10", ((i / 3) % 12) + 1);

            values
                .set(ApiKey::LeftDate(i), date.clone())
                .set(ApiKey::LeftAmountEuro(i), base_amount)
                .set(ApiKey::LeftAmountLocal(i), base_amount)
                .set(ApiKey::RightDate(i), date)
                .set(ApiKey::RightAmountEuro(i), base_amount)
                .set(ApiKey::RightAmountLocal(i), base_amount);
        }

        // === DYNAMISCHE WERTE ===

        // Kategorie 1: 10 Positionen
        for pos in 1..=10u16 {
            values.set_position_row(
                1,
                pos,
                format!("Personal {}", pos),
                30000.0 - (pos as f64 - 1.0) * 1500.0,
                15000.0 - (pos as f64 - 1.0) * 750.0,
                15000.0 - (pos as f64 - 1.0) * 750.0,
                "",
            );
        }

        // Kategorie 2: 12 Positionen
        for pos in 1..=12u16 {
            values.set_position_row(
                2,
                pos,
                format!("Sachkosten {}", pos),
                6000.0 - (pos as f64 - 1.0) * 300.0,
                3000.0 - (pos as f64 - 1.0) * 150.0,
                3000.0 - (pos as f64 - 1.0) * 150.0,
                "",
            );
        }

        // Kategorie 3: 6 Positionen
        for pos in 1..=6u16 {
            values.set_position_row(
                3,
                pos,
                format!("Reisen {}", pos),
                4000.0 - (pos as f64 - 1.0) * 400.0,
                2000.0 - (pos as f64 - 1.0) * 200.0,
                2000.0 - (pos as f64 - 1.0) * 200.0,
                "",
            );
        }

        // Kategorie 4: 4 Positionen
        for pos in 1..=4u16 {
            values.set_position_row(
                4,
                pos,
                format!("Investition {}", pos),
                10000.0 - (pos as f64 - 1.0) * 1500.0,
                5000.0 - (pos as f64 - 1.0) * 750.0,
                5000.0 - (pos as f64 - 1.0) * 750.0,
                "",
            );
        }

        // Kategorie 5: 3 Positionen
        for pos in 1..=3u16 {
            values.set_position_row(
                5,
                pos,
                format!("Sonstiges {}", pos),
                3000.0 - (pos as f64 - 1.0) * 600.0,
                1500.0 - (pos as f64 - 1.0) * 300.0,
                1500.0 - (pos as f64 - 1.0) * 300.0,
                "",
            );
        }

        // Kategorien 6-7: Header-Input Mode
        values.set_header_input(6, 8000.0, 4000.0, 4000.0, "Raummiete und Nebenkosten");
        values.set_header_input(7, 5000.0, 2500.0, 2500.0, "Sicherheit und Versicherung");

        // Kategorie 8 (Reserve): 4 Positionen (seltene Ausnahme)
        values.set_position_row(
            8,
            1,
            "Unvorhergesehene Notfallausgaben",
            3000.0,
            1500.0,
            1500.0,
            "Naturkatastrophen, Notfallreparaturen",
        );
        values.set_position_row(
            8,
            2,
            "Preisanpassungen Material",
            2000.0,
            1000.0,
            1000.0,
            "Inflationsausgleich",
        );
        values.set_position_row(
            8,
            3,
            "Zusätzliche Personalkosten",
            2000.0,
            1000.0,
            1000.0,
            "Krankheitsvertretung, Notfalleinsätze",
        );
        values.set_position_row(
            8,
            4,
            "Reservemanagement",
            1000.0,
            500.0,
            500.0,
            "Administrativer Aufwand für Reserven",
        );

        let body_config = BodyConfig::new()
            .with_positions(1, 10)
            .with_positions(2, 12)
            .with_positions(3, 6)
            .with_positions(4, 4)
            .with_positions(5, 3)
            .with_positions(6, 0) // Header-Input
            .with_positions(7, 0) // Header-Input
            .with_positions(8, 4); // 4 Positionen als seltene Reserve-Ausnahme

        write_report_with_body(ws, &styles, "_de", &values, &body_config)
            .expect("Failed to write report");

        ws.protect();

        let path = "tests/output/test_reserve_with_positions.xlsx";
        workbook.save(path).expect("Failed to save workbook");

        println!("✓ Test 5: Reserve mit mehreren Positionen");
        println!("   Datei: {}", path);
        println!("   - Kategorien 1-5: Multiple Positionen");
        println!("   - Kategorien 6-7: Header-Input Mode");
        println!("   - Kategorie 8: 4 Positionen (Reserve-Detail)");
    }
}
