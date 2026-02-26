//! Test: Workbook-Protection Integration
//!
//! Demonstriert die neue `create_protected_report()` API mit Workbook-Schutz

use kmw_fb_rust::{
    create_protected_report, ApiKey, BodyConfig, ReportOptions, ReportStyles, ReportValues,
};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("Testing Workbook-Protection Integration...\n");

    // Test 1: Report ohne Workbook-Protection
    {
        println!("1. Creating report WITHOUT workbook protection...");
        let styles = ReportStyles::new();
        let mut values = ReportValues::new();
        values.set(ApiKey::Language, "deutsch");
        values.set(ApiKey::Currency, "EUR");
        values.set(ApiKey::ProjectNumber, "TEST-001");
        values.set(ApiKey::ProjectTitle, "Test Project");
        values.set(ApiKey::ProjectStart, "01.01.2024");
        values.set(ApiKey::ProjectEnd, "31.12.2024");
        values.set(ApiKey::ReportStart, "01.01.2024");
        values.set(ApiKey::ReportEnd, "31.03.2024");

        let body_config = BodyConfig::default();
        let options = ReportOptions::new();

        create_protected_report(
            "tests/output/test_no_wb_protection.xlsx",
            &styles,
            &values,
            &body_config,
            &options,
        )?;
        println!("   ✓ Created: tests/output/test_no_wb_protection.xlsx");
    }

    // Test 2: Report mit Workbook-Protection
    {
        println!("\n2. Creating report WITH workbook protection...");
        let styles = ReportStyles::new();
        let mut values = ReportValues::new();
        values.set(ApiKey::Language, "deutsch");
        values.set(ApiKey::Currency, "EUR");
        values.set(ApiKey::ProjectNumber, "TEST-002");
        values.set(ApiKey::ProjectTitle, "Protected Project");
        values.set(ApiKey::ProjectStart, "01.01.2024");
        values.set(ApiKey::ProjectEnd, "31.12.2024");
        values.set(ApiKey::ReportStart, "01.01.2024");
        values.set(ApiKey::ReportEnd, "31.03.2024");

        let body_config = BodyConfig::default();
        let options = ReportOptions::new().with_workbook_protection("secret123");

        create_protected_report(
            "tests/output/test_with_wb_protection.xlsx",
            &styles,
            &values,
            &body_config,
            &options,
        )?;
        println!("   ✓ Created: tests/output/test_with_wb_protection.xlsx");
        println!("   ℹ Password: secret123");
    }

    // Test 3: Report mit Sheet + Workbook Protection
    {
        println!("\n3. Creating report WITH sheet AND workbook protection...");
        let styles = ReportStyles::new();
        let mut values = ReportValues::new();
        values.set(ApiKey::Language, "deutsch");
        values.set(ApiKey::Currency, "EUR");
        values.set(ApiKey::ProjectNumber, "TEST-003");
        values.set(ApiKey::ProjectTitle, "Fully Protected Project");
        values.set(ApiKey::ProjectStart, "01.01.2024");
        values.set(ApiKey::ProjectEnd, "31.12.2024");
        values.set(ApiKey::ReportStart, "01.01.2024");
        values.set(ApiKey::ReportEnd, "31.03.2024");

        let body_config = BodyConfig::default();
        let options = ReportOptions::with_default_protection()
            .with_workbook_protection("test456")
            .with_hidden_columns_qv();

        create_protected_report(
            "tests/output/test_full_protection.xlsx",
            &styles,
            &values,
            &body_config,
            &options,
        )?;
        println!("   ✓ Created: tests/output/test_full_protection.xlsx");
        println!("   ℹ Sheet Password: (default) - see PROTECTION_DEFAULTS");
        println!("   ℹ Workbook Password: test456");
    }

    println!("\n✅ All tests completed successfully!");
    println!("\nManual verification:");
    println!("1. Open test_no_wb_protection.xlsx -> should allow sheet operations");
    println!(
        "2. Open test_with_wb_protection.xlsx -> try to move/delete sheets (requires password)"
    );
    println!(
        "3. Open test_full_protection.xlsx -> both sheet cells and workbook structure protected"
    );

    Ok(())
}
