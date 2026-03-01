//! KMW-FB Rust - Financial Report Generator
//!
//! High-performance Excel financial report generator with dynamic cost positions,
//! formulas, formatting, and multi-language support.
//!
//! ## Usage
//!
//! ```ignore
//! use kmw_fb_rust::ReportConfig;
//!
//! let config = ReportConfig {
//!     language: "deutsch".into(),
//!     currency: "EUR".into(),
//!     locked: true,
//!     ..ReportConfig::default()
//! };
//! config.write_to("report.xlsx")?;
//! ```
//!
//! ### Batch mit vorberechnetem Hash
//!
//! ```ignore
//! use kmw_fb_rust::{precompute_hash, ReportConfig};
//!
//! let hash = precompute_hash("passwort");
//! for config in &configs {
//!     config.write_to_precomputed("output.xlsx", &hash)?;
//! }
//! ```

pub mod config;
pub mod error;
pub mod lang;
pub mod report;
pub mod workbook_protection;

pub use error::ReportError;
pub use workbook_protection::{
    precompute_hash, precompute_hash_with_spin_count,
    protect_workbook, protect_workbook_precomputed, protect_workbook_with_spin_count,
    PrecomputedHash, ProtectionError, WorkbookProtection,
};

// Public API re-exports
pub use config::{PanelEntry, PositionEntry, ReportConfig, TableEntry};
pub use lang::{
    build_sheet as build_language_sheet,
    build_sheet_with_visibility as build_language_sheet_with_visibility,
};
pub use report::{
    api::{ApiKey, Category, Currency, DateError, Language, ReportDate, ReportValues},
    body::BodyConfig,
    options::{
        FieldValidation, HiddenRanges, ReportOptions, RowGroup, RowGrouping,
        SheetProtection, ValidationError, ValidationErrorStyle, ValidationRule, ValidationRuleType,
        ValidationTarget,
    },
    writer::write_report_with_options,
};
