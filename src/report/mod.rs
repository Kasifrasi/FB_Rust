//! Report generation engine.
//!
//! ## Module Structure
//!
//! - `api`: Public API types (ApiKey, Language, Currency, ReportValues, etc.)
//! - `core`: Formula engine (CellRegistry, CellAddr, topological evaluation)
//! - `options`: Report options (ReportOptions, SheetProtection, Validation, etc.)
//! - `styles`: Internal styles and format matrix (ReportStyles, FormatMatrix)
//! - `writer`: Excel writer (write_report_with_options, create_protected_report)
//! - `body`: Dynamic cost position sections

pub mod api;
pub mod body;
pub mod core;
pub mod options;
pub(crate) mod styles;
pub mod writer;

// Re-exports for convenient access
pub use api::{
    register_all_api_cells, ApiKey, Category, CellValue, Currency, DateError, FooterField,
    Language, PositionField, ReportDate, ReportValues,
};
pub use body::{BodyConfig, BodyLayout, CategoryMode};
pub use options::{
    FieldValidation, HiddenRanges, ReportOptions, RowGroup, RowGrouping, SheetProtection,
    ValidationError, ValidationErrorStyle, ValidationRule, ValidationRuleType, ValidationTarget,
};
pub use writer::{
    apply_report_options, setup_freeze_panes, setup_sheet, write_report_with_options, BodyResult,
    MergeRange,
};
