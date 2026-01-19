//! Report Module - Finanzbericht Generator
//!
//! Struktur:
//! - cells: Zell-Adressen und InputCell Definitionen
//! - formulas: Formeldefinitionen und Auswertung
//! - values: Eingabewerte (ReportValues)
//! - formats: Styles und FormatMatrix
//! - layout: Spaltenbreiten, Merges, Freeze Panes
//! - sections: Header, Table, Panel Schreiblogik
//! - writer: Hauptkoordinator

pub mod cells;
pub mod formats;
pub mod formulas;
pub mod layout;
pub mod sections;
pub mod values;
pub mod writer;

// Re-exports für einfachen Zugriff
pub use cells::{CellAddress, HeaderInputCell, InputCell, RightPanelInputCell, TableInputCell};
pub use formats::{build_format_matrix, FormatMatrix, ReportStyles, SectionStyles};
pub use formulas::{
    evaluate_formula, FormulaCache, FormulaDefinition, FormulaType, HEADER_FORMULAS,
};
pub use layout::{setup_default_column_format, setup_freeze_panes, setup_sheet, MergeRange};
pub use values::{CellValue, ReportValues};
pub use writer::write_report;
