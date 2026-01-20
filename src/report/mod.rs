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
//! - `writer`: Main entry points

pub mod api;
pub mod body;
pub mod cells;
pub mod definitions;
pub mod dynamic;
pub mod formats;
pub mod formulas;
pub mod layout;
pub mod registry;
pub mod sections;
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
pub use registry::{
    CellAddr, CellEvaluator, CellKind, CellRegistry, EvalContext, FormulaCell, FormulaCellDeps,
    FormulaDeps, Inputs, RegistryError, SheetRef, Sheets, StaticVal, Statics,
};
pub use values::{CellValue, ReportValues};
pub use writer::{write_report, write_report_with_body};
