//! Dynamischer Body-Bereich (Kostenkategorien)
//!
//! Enthält:
//! - `BodyConfig`: Konfiguration (Anzahl Positionen pro Kategorie)
//! - `BodyLayout`: Berechnetes Layout (Zeilen-Positionen)
//! - `write_body_structure`: Schreibt die Struktur (Labels, Blanks, Formeln)

mod config;
mod layout;
mod writer;

pub use config::BodyConfig;
pub use layout::{BodyLayout, CategoryLayout, PositionRange};
pub use writer::{write_body_structure, BodyResult};
