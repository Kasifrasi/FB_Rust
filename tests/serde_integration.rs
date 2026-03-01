//! Serde Integration Tests
//!
//! Tests JSON serialization/deserialization of all public API types
//! from an external user perspective (via lib.rs re-exports).
//!
//! Run: cargo test --features serde --test serde_integration

#![cfg(feature = "serde")]

use fb_rust::{
    BodyConfig, Category, Currency, Language, PanelEntry, PositionEntry, ReportBody, ReportConfig,
    ReportDate, ReportFooter, ReportHeader, ReportOptions, RowGroup, RowGrouping, TableEntry,
};
use std::collections::HashMap;

// ============================================================================
// Group 1: ReportConfig — Full Roundtrip
// ============================================================================

#[test]
fn test_report_config_full_roundtrip() {
    let config = ReportConfig {
        header: ReportHeader {
            language: Language::English,
            currency: Currency::usd(),
            project_number: Some("PROJ-2025-001".to_string()),
            project_title: Some("Bildungsprojekt".to_string()),
            project_start: Some("01.01.2025".to_string()),
            project_end: Some("31.12.2025".to_string()),
            report_start: Some("01.01.2025".to_string()),
            report_end: Some("30.06.2025".to_string()),
        },
        body: ReportBody {
            table: vec![TableEntry {
                index: 0,
                approved_budget: Some(50000.0),
                income_report: Some(25000.0),
                income_total: Some(25000.0),
                reason: Some("Spende".to_string()),
            }],
            left_panel: vec![PanelEntry {
                index: 0,
                date: Some("15.01.2025".to_string()),
                amount_euro: Some(1000.0),
                amount_local: Some(1100.0),
            }],
            right_panel: vec![PanelEntry {
                index: 0,
                date: Some("20.02.2025".to_string()),
                amount_euro: Some(500.0),
                amount_local: None,
            }],
            positions: vec![
                PositionEntry {
                    category: 1,
                    position: 1,
                    description: Some("Personalkosten".to_string()),
                    approved: Some(18000.0),
                    income_report: Some(9000.0),
                    income_total: Some(9000.0),
                    remark: None,
                },
                PositionEntry {
                    category: 6,
                    position: 0,
                    description: None,
                    approved: Some(3000.0),
                    income_report: Some(1500.0),
                    income_total: Some(1500.0),
                    remark: Some("Verwaltung".to_string()),
                },
            ],
            body_positions: [(1u8, 10u16), (2, 5), (6, 0)].into_iter().collect(),
        },
        footer: ReportFooter {
            bank: Some(8500.0),
            kasse: Some(250.50),
            sonstiges: None,
        },
        options: ReportOptions {
            sheet_password: Some("blatt_pw".to_string()),
            workbook_password: Some("geheim".to_string()),
            hide_columns_qv: true,
            hide_language_sheet: true,
            row_grouping: Some(
                RowGrouping::new()
                    .add_group(10, 20)
                    .add_collapsed_group(25, 30)
                    .with_symbols_above(true),
            ),
        },
    };

    let json = serde_json::to_string(&config).expect("serialize");
    let deserialized: ReportConfig = serde_json::from_str(&json).expect("deserialize");

    // Header
    assert_eq!(deserialized.header.language, Language::English);
    assert_eq!(deserialized.header.currency, Currency::usd());
    assert_eq!(
        deserialized.header.project_number.as_deref(),
        Some("PROJ-2025-001")
    );
    assert_eq!(
        deserialized.header.project_title.as_deref(),
        Some("Bildungsprojekt")
    );
    assert_eq!(
        deserialized.header.project_start.as_deref(),
        Some("01.01.2025")
    );
    assert_eq!(
        deserialized.header.project_end.as_deref(),
        Some("31.12.2025")
    );
    assert_eq!(
        deserialized.header.report_start.as_deref(),
        Some("01.01.2025")
    );
    assert_eq!(
        deserialized.header.report_end.as_deref(),
        Some("30.06.2025")
    );

    // Table
    assert_eq!(deserialized.body.table.len(), 1);
    assert_eq!(deserialized.body.table[0].index, 0);
    assert_eq!(deserialized.body.table[0].approved_budget, Some(50000.0));
    assert_eq!(
        deserialized.body.table[0].reason.as_deref(),
        Some("Spende")
    );

    // Panels
    assert_eq!(deserialized.body.left_panel.len(), 1);
    assert_eq!(deserialized.body.left_panel[0].amount_euro, Some(1000.0));
    assert_eq!(deserialized.body.right_panel.len(), 1);
    assert_eq!(deserialized.body.right_panel[0].amount_local, None);

    // Positions
    assert_eq!(deserialized.body.positions.len(), 2);
    assert_eq!(deserialized.body.positions[0].category, 1);
    assert_eq!(deserialized.body.positions[0].position, 1);
    assert_eq!(
        deserialized.body.positions[0].description.as_deref(),
        Some("Personalkosten")
    );
    assert_eq!(deserialized.body.positions[1].position, 0);

    // Body positions
    assert_eq!(deserialized.body.body_positions[&1], 10);
    assert_eq!(deserialized.body.body_positions[&2], 5);
    assert_eq!(deserialized.body.body_positions[&6], 0);

    // Footer
    assert_eq!(deserialized.footer.bank, Some(8500.0));
    assert_eq!(deserialized.footer.kasse, Some(250.50));
    assert_eq!(deserialized.footer.sonstiges, None);

    // Options
    assert_eq!(
        deserialized.options.workbook_password.as_deref(),
        Some("geheim")
    );
    assert!(deserialized.options.hide_columns_qv);
    assert!(deserialized.options.hide_language_sheet);

    // Row Grouping
    let rg = deserialized.options.row_grouping.as_ref().unwrap();
    assert!(rg.symbols_above);
    assert_eq!(rg.groups().len(), 2);
    assert_eq!(rg.groups()[0].start_row, 10);
    assert_eq!(rg.groups()[0].end_row, 20);
    assert!(!rg.groups()[0].collapsed);
    assert_eq!(rg.groups()[1].start_row, 25);
    assert!(rg.groups()[1].collapsed);
}

#[test]
fn test_report_config_minimal_roundtrip() {
    let config = ReportConfig::default();

    let json = serde_json::to_string(&config).expect("serialize");
    let deserialized: ReportConfig = serde_json::from_str(&json).expect("deserialize");

    assert_eq!(deserialized.header.language, Language::Deutsch);
    assert_eq!(deserialized.header.currency, Currency::eur());
    assert!(deserialized.options.sheet_password.is_none());
    assert!(deserialized.body.table.is_empty());
    assert!(deserialized.body.left_panel.is_empty());
    assert!(deserialized.body.right_panel.is_empty());
    assert!(deserialized.body.positions.is_empty());
    assert!(deserialized.header.project_number.is_none());
    assert!(deserialized.footer.bank.is_none());
    assert!(deserialized.options.row_grouping.is_none());
}

#[test]
fn test_report_config_empty_json() {
    // All sub-structs have serde(default), so {} should work
    let json = r#"{}"#;
    let config: ReportConfig = serde_json::from_str(json).expect("deserialize empty");

    assert_eq!(config.header.language, Language::Deutsch);
    assert_eq!(config.header.currency, Currency::eur());
    assert!(config.options.sheet_password.is_none());
    assert!(config.body.table.is_empty());
}

#[test]
fn test_report_config_partial_json() {
    // Only header provided, everything else defaults
    let json = r#"{
        "header": {
            "language": "english",
            "currency": "GBP"
        }
    }"#;

    let config: ReportConfig = serde_json::from_str(json).expect("deserialize partial");

    assert_eq!(config.header.language, Language::English);
    assert_eq!(config.header.currency, Currency::gbp());
    assert!(config.header.project_number.is_none());
    assert!(config.body.table.is_empty());
    assert!(config.footer.bank.is_none());
    assert!(config.options.sheet_password.is_none());
}

// ============================================================================
// Group 2: Language — rename_all="lowercase"
// ============================================================================

#[test]
fn test_language_serialize() {
    let cases: &[(Language, &str)] = &[
        (Language::Deutsch, r#""deutsch""#),
        (Language::English, r#""english""#),
        (Language::Francais, r#""francais""#),
        (Language::Espanol, r#""espanol""#),
        (Language::Portugues, r#""portugues""#),
    ];

    for (lang, expected) in cases {
        let json = serde_json::to_string(lang).expect("serialize language");
        assert_eq!(
            &json, expected,
            "Language::{:?} should serialize to {}",
            lang, expected
        );
    }
}

#[test]
fn test_language_deserialize() {
    let cases: &[(&str, Language)] = &[
        (r#""deutsch""#, Language::Deutsch),
        (r#""english""#, Language::English),
        (r#""francais""#, Language::Francais),
        (r#""espanol""#, Language::Espanol),
        (r#""portugues""#, Language::Portugues),
    ];

    for (json, expected) in cases {
        let lang: Language = serde_json::from_str(json).expect("deserialize language");
        assert_eq!(lang, *expected);
    }
}

#[test]
fn test_language_deserialize_invalid() {
    let invalid = [r#""german""#, r#""Deutsch""#, r#""de""#, r#""ENGLISH""#];

    for json in &invalid {
        let result = serde_json::from_str::<Language>(json);
        assert!(result.is_err(), "{} should fail to deserialize", json);
    }
}

// ============================================================================
// Group 3: Currency — try_from validated
// ============================================================================

#[test]
fn test_currency_serde_roundtrip() {
    let eur = Currency::eur();
    let json = serde_json::to_string(&eur).expect("serialize");
    assert_eq!(json, r#""EUR""#);

    let usd: Currency = serde_json::from_str(r#""USD""#).expect("deserialize");
    assert_eq!(usd, Currency::usd());

    let chf = Currency::chf();
    let json = serde_json::to_string(&chf).expect("serialize");
    let back: Currency = serde_json::from_str(&json).expect("deserialize");
    assert_eq!(back, chf);
}

#[test]
fn test_currency_deserialize_invalid() {
    let invalid = [r#""INVALID""#, r#""XYZ""#, r#"""#];

    for json in &invalid {
        let result = serde_json::from_str::<Currency>(json);
        assert!(result.is_err(), "{} should fail to deserialize", json);
    }
}

#[test]
fn test_currency_deserialize_case_insensitive() {
    let lower: Currency = serde_json::from_str(r#""eur""#).expect("lowercase");
    assert_eq!(lower.as_str(), "EUR");
}

// ============================================================================
// Group 4: Category — Standard Enum
// ============================================================================

#[test]
fn test_category_serde_roundtrip() {
    let all_categories = Category::all();

    let expected_names = [
        "Bauausgaben",
        "Investitionen",
        "Personalausgaben",
        "Projektaktivitaeten",
        "Projektverwaltung",
        "Evaluierung",
        "Audit",
        "Reserve",
    ];

    for (cat, expected_name) in all_categories.iter().zip(expected_names.iter()) {
        let json = serde_json::to_string(cat).expect("serialize category");
        assert_eq!(json, format!("\"{}\"", expected_name));

        let back: Category = serde_json::from_str(&json).expect("deserialize category");
        assert_eq!(*cat, back);
    }
}

// ============================================================================
// Group 5: ReportDate — try_from/into String
// ============================================================================

#[test]
fn test_report_date_serde_iso_roundtrip() {
    let date = ReportDate::new(2024, 6, 15).unwrap();

    let json = serde_json::to_string(&date).expect("serialize");
    assert_eq!(json, r#""2024-06-15""#);

    let back: ReportDate = serde_json::from_str(&json).expect("deserialize");
    assert_eq!(back.year(), 2024);
    assert_eq!(back.month(), 6);
    assert_eq!(back.day(), 15);
}

#[test]
fn test_report_date_deserialize_formats() {
    let iso: ReportDate = serde_json::from_str(r#""2024-06-15""#).expect("ISO");
    assert_eq!(iso.year(), 2024);
    assert_eq!(iso.month(), 6);
    assert_eq!(iso.day(), 15);

    let de: ReportDate = serde_json::from_str(r#""15.06.2024""#).expect("DE");
    assert_eq!(de.day(), 15);
    assert_eq!(de.month(), 6);
    assert_eq!(de.year(), 2024);

    let slash: ReportDate = serde_json::from_str(r#""15/06/2024""#).expect("slash");
    assert_eq!(slash.day(), 15);
    assert_eq!(slash.month(), 6);
}

#[test]
fn test_report_date_deserialize_invalid() {
    assert!(serde_json::from_str::<ReportDate>(r#""invalid""#).is_err());
    assert!(serde_json::from_str::<ReportDate>(r#""2024-13-01""#).is_err());
    assert!(serde_json::from_str::<ReportDate>(r#""2024-02-30""#).is_err());
}

// ============================================================================
// Group 6: BodyConfig — from/into HashMap
// ============================================================================

#[test]
fn test_body_config_serializes_as_hashmap() {
    let config = BodyConfig::new().with_positions(1, 10).with_positions(6, 0);

    let json = serde_json::to_string(&config).expect("serialize");
    let map: HashMap<String, u16> = serde_json::from_str(&json).expect("as raw map");

    assert_eq!(map["1"], 10);
    assert_eq!(map["6"], 0);
    assert!(!map.contains_key("positions"));
}

#[test]
fn test_body_config_deserializes_from_hashmap() {
    let json = r#"{"1": 20, "2": 20, "3": 30}"#;
    let config: BodyConfig = serde_json::from_str(json).expect("deserialize");

    assert_eq!(config.position_count(1), 20);
    assert_eq!(config.position_count(2), 20);
    assert_eq!(config.position_count(3), 30);
    assert_eq!(config.position_count(4), 0);
}

#[test]
fn test_body_config_roundtrip() {
    let original = BodyConfig::new()
        .with_positions(1, 5)
        .with_positions(2, 10)
        .with_positions(6, 0)
        .with_positions(8, 0);

    let json = serde_json::to_string(&original).expect("serialize");
    let back: BodyConfig = serde_json::from_str(&json).expect("deserialize");

    assert_eq!(back.position_count(1), 5);
    assert_eq!(back.position_count(2), 10);
    assert_eq!(back.position_count(6), 0);
    assert_eq!(back.position_count(8), 0);
}

// ============================================================================
// Group 7: Helper types
// ============================================================================

#[test]
fn test_table_entry_roundtrip() {
    let entry = TableEntry {
        index: 2,
        approved_budget: Some(100000.0),
        income_report: Some(50000.0),
        income_total: Some(50000.0),
        reason: Some("Spenden".to_string()),
    };

    let json = serde_json::to_string(&entry).expect("serialize");
    let back: TableEntry = serde_json::from_str(&json).expect("deserialize");

    assert_eq!(back.index, 2);
    assert_eq!(back.approved_budget, Some(100000.0));
    assert_eq!(back.reason.as_deref(), Some("Spenden"));

    let entry_null = TableEntry {
        index: 0,
        approved_budget: None,
        income_report: None,
        income_total: None,
        reason: None,
    };

    let json = serde_json::to_string(&entry_null).expect("serialize null");
    let back: TableEntry = serde_json::from_str(&json).expect("deserialize null");
    assert_eq!(back.index, 0);
    assert!(back.approved_budget.is_none());
    assert!(back.reason.is_none());
}

#[test]
fn test_panel_entry_roundtrip() {
    let entry = PanelEntry {
        index: 5,
        date: Some("15.03.2025".to_string()),
        amount_euro: Some(750.25),
        amount_local: None,
    };

    let json = serde_json::to_string(&entry).expect("serialize");
    let back: PanelEntry = serde_json::from_str(&json).expect("deserialize");

    assert_eq!(back.index, 5);
    assert_eq!(back.date.as_deref(), Some("15.03.2025"));
    assert_eq!(back.amount_euro, Some(750.25));
    assert!(back.amount_local.is_none());
}

#[test]
fn test_position_entry_roundtrip() {
    let pos = PositionEntry {
        category: 1,
        position: 3,
        description: Some("Reisekosten".to_string()),
        approved: Some(5000.0),
        income_report: Some(2500.0),
        income_total: Some(2500.0),
        remark: Some("Dienstreise".to_string()),
    };

    let json = serde_json::to_string(&pos).expect("serialize");
    let back: PositionEntry = serde_json::from_str(&json).expect("deserialize");
    assert_eq!(back.category, 1);
    assert_eq!(back.position, 3);
    assert_eq!(back.description.as_deref(), Some("Reisekosten"));

    let header = PositionEntry {
        category: 6,
        position: 0,
        description: None,
        approved: Some(3000.0),
        income_report: Some(1500.0),
        income_total: Some(1500.0),
        remark: None,
    };

    let json = serde_json::to_string(&header).expect("serialize header");
    let back: PositionEntry = serde_json::from_str(&json).expect("deserialize header");
    assert_eq!(back.category, 6);
    assert_eq!(back.position, 0);
    assert!(back.description.is_none());
}

// ============================================================================
// Group 8: RowGrouping / RowGroup
// ============================================================================

#[test]
fn test_row_grouping_roundtrip() {
    let grouping = RowGrouping::new()
        .add_group(10, 20)
        .add_collapsed_group(25, 30)
        .with_symbols_above(true);

    let json = serde_json::to_string(&grouping).expect("serialize");
    let back: RowGrouping = serde_json::from_str(&json).expect("deserialize");

    assert!(back.symbols_above);
    assert_eq!(back.groups().len(), 2);

    assert_eq!(back.groups()[0].start_row, 10);
    assert_eq!(back.groups()[0].end_row, 20);
    assert!(!back.groups()[0].collapsed);

    assert_eq!(back.groups()[1].start_row, 25);
    assert_eq!(back.groups()[1].end_row, 30);
    assert!(back.groups()[1].collapsed);
}

#[test]
fn test_row_group_from_json() {
    let json = r#"{"start_row": 5, "end_row": 15, "collapsed": true}"#;
    let group: RowGroup = serde_json::from_str(json).expect("deserialize");

    assert_eq!(group.start_row, 5);
    assert_eq!(group.end_row, 15);
    assert!(group.collapsed);
}

// ============================================================================
// Group 9: ReportConfig from Tauri-typical JSON (nested format)
// ============================================================================

#[test]
fn test_report_config_from_typescript_json() {
    let json = r#"{
        "header": {
            "language": "deutsch",
            "currency": "EUR",
            "project_number": "2025-001",
            "project_title": "Klimaschutzprojekt"
        },
        "body": {
            "body_positions": { "1": 10, "2": 5 },
            "positions": [
                {
                    "category": 1,
                    "position": 1,
                    "description": "Baukosten",
                    "approved": 50000.0,
                    "income_report": null,
                    "income_total": null,
                    "remark": null
                }
            ]
        },
        "footer": {
            "bank": 12000.0,
            "kasse": 500.0
        },
        "options": {
            "sheet_password": "geheim",
            "hide_columns_qv": false,
            "hide_language_sheet": false
        }
    }"#;

    let config: ReportConfig = serde_json::from_str(json).expect("deserialize TypeScript JSON");

    assert_eq!(config.header.language, Language::Deutsch);
    assert_eq!(config.header.currency, Currency::eur());
    assert_eq!(config.header.project_number.as_deref(), Some("2025-001"));
    assert_eq!(
        config.header.project_title.as_deref(),
        Some("Klimaschutzprojekt")
    );
    assert_eq!(config.options.sheet_password.as_deref(), Some("geheim"));
    assert_eq!(config.body.body_positions[&1], 10);
    assert_eq!(config.body.body_positions[&2], 5);
    assert_eq!(config.body.positions.len(), 1);
    assert_eq!(config.body.positions[0].approved, Some(50000.0));
    assert_eq!(
        config.body.positions[0].description.as_deref(),
        Some("Baukosten")
    );
    assert_eq!(config.footer.bank, Some(12000.0));
    assert_eq!(config.footer.kasse, Some(500.0));
    assert!(config.footer.sonstiges.is_none());
    // serde(default) fields that are missing → empty
    assert!(config.body.table.is_empty());
    assert!(config.body.left_panel.is_empty());
    assert!(config.body.right_panel.is_empty());
    assert!(config.options.row_grouping.is_none());
}

// ============================================================================
// Group 10: Edge Cases
// ============================================================================

#[test]
fn test_report_config_unknown_fields_rejected() {
    // deny_unknown_fields on top level
    let json = r#"{ "unknown_field": true }"#;
    let result = serde_json::from_str::<ReportConfig>(json);
    assert!(result.is_err(), "unknown top-level fields should be rejected");

    // deny_unknown_fields on sub-structs
    let json = r#"{ "header": { "language": "deutsch", "currency": "EUR", "unknown": true } }"#;
    let result = serde_json::from_str::<ReportConfig>(json);
    assert!(
        result.is_err(),
        "unknown header fields should be rejected"
    );
}

#[test]
fn test_body_positions_empty_map() {
    let json = r#"{
        "body": { "body_positions": {} }
    }"#;

    let config: ReportConfig = serde_json::from_str(json).expect("empty body_positions");
    assert!(config.body.body_positions.is_empty());
}

#[test]
fn test_body_config_empty_hashmap() {
    let json = r#"{}"#;
    let config: BodyConfig = serde_json::from_str(json).expect("empty map");
    assert_eq!(config.position_count(1), 0);
}

#[test]
fn test_report_config_all_null_optionals() {
    let json = r#"{
        "header": {
            "language": "english",
            "currency": "GBP",
            "project_number": null,
            "project_title": null,
            "project_start": null,
            "project_end": null,
            "report_start": null,
            "report_end": null
        },
        "body": {
            "body_positions": {"1": 5}
        },
        "footer": {
            "bank": null,
            "kasse": null,
            "sonstiges": null
        },
        "options": {
            "workbook_password": null,
            "hide_columns_qv": false,
            "hide_language_sheet": false
        }
    }"#;

    let config: ReportConfig = serde_json::from_str(json).expect("all-null optionals");

    assert_eq!(config.header.language, Language::English);
    assert_eq!(config.header.currency, Currency::gbp());
    assert!(config.header.project_number.is_none());
    assert!(config.header.project_title.is_none());
    assert!(config.footer.bank.is_none());
    assert!(config.options.workbook_password.is_none());
    assert_eq!(config.body.body_positions[&1], 5);
}
