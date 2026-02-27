//! Excel-Writer für den Finanzbericht
//!
//! Dieses Modul enthält:
//! - `main`: Haupt-Schreibfunktionen (write_report_with_options)
//! - `layout`: Sheet-Setup (Spaltenbreiten, Freeze Panes)
//! - `structure`: Report-Struktur (Merges, Blanks, Strings, Validierungen)

mod layout;
mod main;
mod structure;

pub use layout::{setup_freeze_panes, setup_sheet, MergeRange};
pub use main::{apply_report_options, create_protected_report, write_report_with_options, BodyResult};
