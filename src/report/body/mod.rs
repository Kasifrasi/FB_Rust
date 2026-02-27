//! Dynamischer Body-Bereich (Kostenkategorien) — Domänen-Logik
//!
//! Enthält:
//! - `BodyConfig`: Konfiguration (Anzahl Positionen pro Kategorie)
//! - `BodyLayout`: Berechnetes Layout (Zeilen-Positionen)
//! - `register_body_formulas`: Registriert Body-Formeln mit der einheitlichen Registry
//! - `register_footer_formulas`: Registriert Footer-Formeln mit der einheitlichen Registry
//!
//! Schreib-Logik → `writer/sections/body.rs` und `writer/sections/footer.rs`

mod config;
mod formulas;
mod layout;

pub use config::BodyConfig;
pub use formulas::{register_body_formulas, register_footer_formulas};
pub use layout::{
    BodyLayout, CategoryLayout, CategoryMode, FooterLayout, PositionRange, TOTAL_LABEL_INDEX,
};
