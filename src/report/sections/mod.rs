//! Sections des Finanzberichts
//!
//! Jede Section ist für einen bestimmten Bereich verantwortlich:
//! - header: Zeilen 0-9 (Titel, Projektinfos, etc.)
//! - table: Zeilen 10-19 (Einnahmen-Tabelle)
//! - panel: Zeilen 10-30, Spalten J-V (Right Panel)
//! - prebody: Zeilen 20-25 (Spacer + Footer-Tabelle vor Body)

pub mod header;
pub mod panel;
pub mod prebody;
pub mod table;

pub use header::write_header_section;
pub use panel::write_panel_section;
pub use prebody::{write_prebody_section, write_prebody_section_unified};
pub use table::write_table_section;
