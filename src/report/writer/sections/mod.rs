//! Report-Sections (statisch + dynamisch)
//!
//! Diese Module schreiben alle Bereiche des Reports:
//! - `header`: Kopfbereich (Zeilen 1-9)
//! - `table`: Einnahmen-Tabelle (Zeilen 14-20)
//! - `panel`: Right Panel (Zeilen 14-31)
//! - `prebody`: Prebody-Bereich (Zeilen 22-26)
//! - `body`: Dynamischer Body-Bereich (Kostenkategorien)
//! - `footer`: Footer-Bereich nach dem Body

mod body;
mod footer;
mod header;
mod panel;
mod prebody;
mod table;

pub use body::{write_body_structure_with_values, BodyResult};
pub use crate::report::body::FooterLayout;
pub use footer::{write_footer, write_footer_values};
pub use header::write_header_section;
pub use panel::write_panel_section;
pub use prebody::write_prebody_section;
pub use table::write_table_section;
