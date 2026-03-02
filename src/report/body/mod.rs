//! Dynamischer Body-Bereich (Kostenkategorien) — Domänen-Logik
//!
//! Enthält:
//! - `BodyConfig`: Konfiguration (Anzahl Positionen pro Kategorie)
//! - `BodyLayout`: Berechnetes Layout (Zeilen-Positionen)
//! - `register_body_formulas`: Registriert Body-Formeln mit der einheitlichen Registry
//! - `register_footer_formulas`: Registriert Footer-Formeln mit der einheitlichen Registry
//!
//! Schreib-Logik → `writer/structure.rs` (Merges/Blanks) und `writer/main.rs` (Zellwerte)

pub(crate) mod config;
mod layout;
mod registration;

pub use config::BodyConfig;
pub use registration::{register_body_formulas, register_footer_formulas};
pub use layout::{
    BodyLayout, CategoryLayout, CategoryMode, FooterLayout, PositionRange, TOTAL_LABEL_INDEX,
};
