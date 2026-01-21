//! Dynamischer Body-Bereich (Kostenkategorien) und Footer
//!
//! Enthält:
//! - `BodyConfig`: Konfiguration (Anzahl Positionen pro Kategorie)
//! - `BodyLayout`: Berechnetes Layout (Zeilen-Positionen)
//! - `write_body_structure`: Schreibt die Struktur (Labels, Blanks, Formeln)
//! - `write_body_structure_with_values`: Schreibt die Struktur MIT API-Werten
//! - `FooterLayout`: Layout für den Footer-Bereich
//! - `write_footer`: Schreibt den Footer nach dem Body
//! - `register_body_formulas`: Registriert Body-Formeln mit der einheitlichen Registry
//! - `register_footer_formulas`: Registriert Footer-Formeln mit der einheitlichen Registry

mod config;
mod footer;
mod formulas;
mod layout;
mod writer;

pub use config::BodyConfig;
pub use footer::{write_footer, write_footer_values, FooterLayout};
pub use formulas::{register_body_formulas, register_footer_formulas, DynRegistry};
pub use layout::{BodyLayout, CategoryLayout, CategoryMode, PositionRange};
pub use writer::{write_body_structure, write_body_structure_with_values, BodyResult};
