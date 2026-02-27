//! KMW-FB Rust - Financial Report Generator
//!
//! High-performance Excel financial report generator with dynamic cost positions,
//! formulas, formatting, and multi-language support.
//!
//! ## Features
//!
//! - Dynamic cost position sections with configurable categories
//! - Multi-language support (German, English, French, Spanish, Portuguese)
//! - Registry-based formula evaluation with caching
//! - Sheet protection with configurable options
//! - Data validation for input cells
//!
//! ## Usage
//!
//! ```ignore
//! use kmw_fb_rust::{
//!     BodyConfig, ReportStyles, ReportValues, ReportOptions, SheetProtection,
//!     write_report_with_options,
//! };
//! use rust_xlsxwriter::Workbook;
//!
//! let mut workbook = Workbook::new();
//! let worksheet = workbook.add_worksheet();
//!
//! let styles = ReportStyles::new();
//! let values = ReportValues::new()
//!     .with_language("deutsch")
//!     .with_currency("EUR");
//! let body_config = BodyConfig::default();
//! let options = ReportOptions::with_default_protection()
//!     .with_hidden_columns_qv();
//!
//! write_report_with_options(worksheet, &styles, "", &values, &body_config, &options)?;
//! workbook.save("report.xlsx")?;
//! ```

pub mod lang;
pub mod report;

// Workbook Protection Module
#[path = "workbook_protection/workbook_protect.rs"]
pub mod workbook_protection;

// Public API re-exports
pub use lang::{
    build_sheet as build_language_sheet,
    build_sheet_with_visibility as build_language_sheet_with_visibility,
};
pub use report::{
    api::{ApiKey, Category, Currency, DateError, Language, ReportDate, ReportValues},
    body::BodyConfig,
    format::{
        FieldValidation, HiddenRanges, ReportOptions, ReportStyles, RowGroup, RowGrouping,
        SheetProtection, ValidationError, ValidationErrorStyle, ValidationRule, ValidationRuleType,
        ValidationTarget, WorkbookProtection,
    },
    writer::{apply_report_options, create_protected_report, write_report_with_options},
};
