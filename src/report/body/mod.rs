//! Dynamischer Body-Bereich (Kostenkategorien) — Domänen-Logik
//!
//! Enthält:
//! - `BodyConfig`: Konfiguration (Anzahl Positionen pro Kategorie)
//! - `BodyLayout`: Berechnetes Layout (Zeilen-Positionen)
//!
//! Schreib-Logik → `writer/structure.rs` (Merges/Blanks) und `writer/main.rs` (Zellwerte)

pub(crate) mod config;
mod layout;

pub use config::BodyConfig;
pub use layout::{
    BodyLayout, CategoryLayout, CategoryMode, FooterLayout, PositionRange, TOTAL_LABEL_INDEX,
};
