//! Statische Report-Sections
//!
//! Diese Module schreiben die statischen Bereiche des Reports:
//! - `header`: Kopfbereich (Zeilen 1-9)
//! - `header_new`: Kopfbereich (Neues zellbasiertes System - Proof of Concept)
//! - `table`: Einnahmen-Tabelle (Zeilen 14-20)
//! - `panel`: Right Panel (Zeilen 14-31)
//! - `prebody`: Prebody-Bereich (Zeilen 22-26)

mod header;
mod header_new;
mod panel;
mod prebody;
mod prebody_new;
mod table;
mod table_new;

pub use header::write_header_section;
pub use header_new::write_header_new;
pub use panel::write_panel_section;
pub use prebody::write_prebody_section;
pub use prebody_new::write_prebody_new;
pub use table::write_table_section;
pub use table_new::write_table_new;
