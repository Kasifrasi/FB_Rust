//! Dynamischer Body-Bereich (Kostenkategorien)
//!
//! Enthält:
//! - `BodyConfig`: Konfiguration (Anzahl Positionen pro Kategorie)
//! - `BodyLayout`: Berechnetes Layout (Zeilen-Positionen)
//! - `write_body_structure`: Schreibt die Struktur (Labels, Blanks, Formeln)
//! - `write_body_structure_with_values`: Schreibt die Struktur MIT API-Werten

mod config;
mod layout;
mod writer;

pub use config::BodyConfig;
pub use layout::{BodyLayout, CategoryLayout, CategoryMode, PositionRange};
pub use writer::{write_body_structure, write_body_structure_with_values, BodyResult};
