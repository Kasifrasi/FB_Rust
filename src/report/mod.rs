//! Report generation engine.
//!
//! ## Module Structure
//!
//! - `api`: Public API types (ApiKey, Language, Currency, ReportValues, etc.)
//! - `core`: Formula engine (CellRegistry, CellAddr, topological evaluation)
//! - `format`: Styles and protection (ReportStyles, FormatMatrix, SheetProtection)
//! - `writer`: Excel writer (write_report_with_options, create_protected_report)
//! - `body`: Dynamic cost position sections

pub mod api;
pub mod body;
pub mod core;
pub mod format;
pub mod writer;

// Re-exports for convenient access
pub use api::{
    register_all_api_cells, ApiKey, Category, CellValue, Currency, DateError, FooterField,
    Language, PositionField, ReportDate, ReportValues,
};
pub use body::{BodyConfig, BodyLayout, CategoryMode};
pub use core::{
    build_registry, CellAddr, CellEvaluator, CellKind, CellRegistry, EvalContext,
    FormulaCell, FormulaCellDeps, FormulaDeps, Inputs,
    RegistryError, SheetRef, Sheets, StaticVal, Statics,
};
pub use format::{
    build_format_matrix, extend_format_matrix_with_body, extend_format_matrix_with_footer,
    extend_format_matrix_with_prebody, BodyStyles, FieldValidation, FormatMatrix, HiddenRange,
    HiddenRanges, NumericRule, ReportOptions, ReportStyles, SectionStyles, SheetProtection,
    ValidationError, ValidationErrorStyle, ValidationRule, ValidationRuleType, ValidationTarget,
};
pub use writer::{
    apply_report_options, create_protected_report, setup_freeze_panes, setup_sheet,
    write_report_with_options, BodyResult, MergeRange,
};
