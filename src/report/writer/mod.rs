//! Excel-Writer für den Finanzbericht
//!
//! Dieses Modul enthält:
//! - `main`: Haupt-Schreibfunktionen (write_report_with_options)
//! - `layout`: Sheet-Setup (Spaltenbreiten, Freeze Panes)
//! - `sections`: Statische Report-Sections

mod layout;
mod main;
pub mod sections;

pub use layout::{setup_freeze_panes, setup_sheet, MergeRange};
pub use main::{apply_report_options, write_report_with_options};
