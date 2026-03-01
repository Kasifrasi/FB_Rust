//! Excel-Writer für den Finanzbericht
//!
//! Dieses Modul enthält:
//! - `engine`: Haupt-Schreibfunktionen (write_report_with_options)
//! - `layout`: Sheet-Setup (Spaltenbreiten, Freeze Panes)
//! - `structure`: Report-Struktur (Merges, Blanks, Strings, Validierungen)

mod engine;
mod layout;
mod structure;

pub use layout::{setup_freeze_panes, setup_sheet, MergeRange};
pub use engine::{apply_sheet_options, write_report_with_options, BodyResult};
pub(crate) use engine::{create_protected_report, create_protected_report_precomputed};
