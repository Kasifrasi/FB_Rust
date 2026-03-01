//! Test: Workbook-Protection Integration
//!
//! Demonstriert die `ReportConfig` API mit verschiedenen Schutz-Stufen

use fb_rust::ReportConfig;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("Testing Workbook-Protection Integration...\n");

    // Test 1: Report ohne Schutz
    {
        println!("1. Creating report WITHOUT workbook protection...");
        let config = ReportConfig {
            language: "deutsch".to_string(),
            currency: "EUR".to_string(),
            project_number: Some("TEST-001".to_string()),
            project_title: Some("Test Project".to_string()),
            project_start: Some("01.01.2024".to_string()),
            project_end: Some("31.12.2024".to_string()),
            report_start: Some("01.01.2024".to_string()),
            report_end: Some("31.03.2024".to_string()),
            ..ReportConfig::default()
        };
        config.write_to("output/test_no_wb_protection.xlsx")?;
        println!("   ✓ Created: output/test_no_wb_protection.xlsx");
    }

    // Test 2: Report mit Workbook-Protection (kein Sheet-Schutz)
    {
        println!("\n2. Creating report WITH workbook protection...");
        let config = ReportConfig {
            language: "deutsch".to_string(),
            currency: "EUR".to_string(),
            project_number: Some("TEST-002".to_string()),
            project_title: Some("Protected Project".to_string()),
            project_start: Some("01.01.2024".to_string()),
            project_end: Some("31.12.2024".to_string()),
            report_start: Some("01.01.2024".to_string()),
            report_end: Some("31.03.2024".to_string()),
            workbook_password: Some("secret123".to_string()),
            ..ReportConfig::default()
        };
        config.write_to("output/test_with_wb_protection.xlsx")?;
        println!("   ✓ Created: output/test_with_wb_protection.xlsx");
        println!("   ℹ Password: secret123");
    }

    // Test 3: Report mit Sheet- + Workbook-Protection
    {
        println!("\n3. Creating report WITH sheet AND workbook protection...");
        let config = ReportConfig {
            language: "deutsch".to_string(),
            currency: "EUR".to_string(),
            project_number: Some("TEST-003".to_string()),
            project_title: Some("Fully Protected Project".to_string()),
            project_start: Some("01.01.2024".to_string()),
            project_end: Some("31.12.2024".to_string()),
            report_start: Some("01.01.2024".to_string()),
            report_end: Some("31.03.2024".to_string()),
            locked: true,
            workbook_password: Some("test456".to_string()),
            hide_columns_qv: true,
            hide_language_sheet: true,
            ..ReportConfig::default()
        };
        config.write_to("output/test_full_protection.xlsx")?;
        println!("   ✓ Created: output/test_full_protection.xlsx");
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
