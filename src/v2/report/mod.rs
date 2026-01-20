//! Report Module - Finanzbericht Generator
//!
//! Struktur:
//! - api: API-Zellen Definition (EINZIGE QUELLE DER WAHRHEIT)
//! - registry: Zentrale Zell-Registry (Api, Static, Formula, Empty)
//! - values: Eingabewerte (ReportValues)
//! - formats: Styles und FormatMatrix
//! - layout: Spaltenbreiten, Merges, Freeze Panes
//! - sections: Header, Table, Panel Schreiblogik
//! - body: Dynamischer Body-Bereich (Kostenkategorien)
//! - writer: Hauptkoordinator

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

// Re-exports für einfachen Zugriff
// API-Zellen (neue zentrale Definitionen)
pub use api::{register_all_api_cells, ApiKey, PositionField, SingleRowField};

// Legacy cells (für Kompatibilität, wird später entfernt)
pub use cells::CellAddress;

pub use body::{BodyConfig, BodyLayout, BodyResult};
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
pub use writer::{write_report_v2, write_report_v2_with_body};
