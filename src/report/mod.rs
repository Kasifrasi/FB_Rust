//! Report generation engine.
//!
//! Modules:
//! - `api`: Cell definitions and API keys
//! - `registry`: Cell registry with topological formula evaluation
//! - `values`: Input value storage
//! - `formats`: Styles and format matrix
//! - `layout`: Column widths, merges, freeze panes
//! - `sections`: Static sections (header, table, panel)
//! - `body`: Dynamic cost position sections
//! - `protection`: Sheet protection and data validation
//! - `types`: Type-safe input types (Language, Currency, Category, Date)
//! - `writer`: Main entry points

pub mod api;
pub mod body;
pub mod cells;
pub mod definitions;
pub mod dynamic;
pub mod formats;
pub mod formulas;
pub mod layout;
pub mod protection;
pub mod registry;
pub mod sections;
pub mod types;
pub mod values;
pub mod writer;

// Public API
pub use api::{register_all_api_cells, ApiKey, PositionField};
pub use body::{BodyConfig, BodyLayout, BodyResult, CategoryMode};
pub use cells::CellAddress;
pub use formats::{
    build_format_matrix, extend_format_matrix_with_body, FormatMatrix, ReportStyles, SectionStyles,
};
pub use formulas::{
    evaluate_formula, FormulaCache, FormulaDefinition, FormulaType, HEADER_FORMULAS,
};
pub use layout::{setup_default_column_format, setup_freeze_panes, setup_sheet, MergeRange};
pub use protection::{
    FieldValidation, HiddenRange, HiddenRanges, NumericRule, ReportOptions, SheetProtection,
    ValidationError, ValidationErrorStyle, ValidationRule, ValidationRuleType, ValidationTarget,
};
pub use registry::{
    CellAddr, CellEvaluator, CellKind, CellRegistry, EvalContext, FormulaCell, FormulaCellDeps,
    FormulaDeps, Inputs, RegistryError, SheetRef, Sheets, StaticVal, Statics,
};
pub use types::{Category, Currency, DateError, Language, ReportDate};
pub use values::{CellValue, ReportValues};
pub use writer::{
    apply_report_options, write_report, write_report_with_body, write_report_with_options,
};
