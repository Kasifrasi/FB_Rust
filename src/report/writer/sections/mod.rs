//! Statische Report-Sections
//!
//! Diese Module schreiben die statischen Bereiche des Reports:
//! - `header`: Kopfbereich (Zeilen 1-9)
//! - `table`: Einnahmen-Tabelle (Zeilen 14-20)
//! - `panel`: Right Panel (Zeilen 14-31)
//! - `prebody`: Prebody-Bereich (Zeilen 22-26)

mod header;
mod panel;
mod prebody;
mod table;

pub use header::write_header_section;
pub use panel::write_panel_section;
pub use prebody::write_prebody_section;
pub use table::write_table_section;
