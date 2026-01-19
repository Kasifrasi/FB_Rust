//! Report Module - Finanzbericht Generator
//!
//! Struktur:
//! - registry: Zentrale Zell-Registry (Api, Static, Formula, Empty)
//! - cells: Zell-Adressen und InputCell Definitionen (legacy)
//! - formulas: Formeldefinitionen und Auswertung (legacy)
//! - values: Eingabewerte (ReportValues)
//! - formats: Styles und FormatMatrix
//! - layout: Spaltenbreiten, Merges, Freeze Panes
//! - sections: Header, Table, Panel Schreiblogik
//! - writer: Hauptkoordinator

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
pub mod writer_v2;

// Re-exports für einfachen Zugriff
pub use cells::{CellAddress, HeaderInputCell, InputCell, RightPanelInputCell, TableInputCell};
pub use formats::{build_format_matrix, FormatMatrix, ReportStyles, SectionStyles};
pub use formulas::{
    evaluate_formula, FormulaCache, FormulaDefinition, FormulaType, HEADER_FORMULAS,
};
pub use layout::{setup_default_column_format, setup_freeze_panes, setup_sheet, MergeRange};
pub use registry::{
    ApiKey, CellAddr, CellEvaluator, CellKind, CellRegistry, EvalContext, FormulaCell,
    FormulaCellDeps, FormulaDeps, Inputs, RegistryError, SheetRef, Sheets, StaticVal, Statics,
};
pub use values::{CellValue, ReportValues};
pub use writer::write_report;
