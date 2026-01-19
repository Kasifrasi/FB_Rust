//! Language Module - Sprachversionen und Übersetzungen
//!
//! Enthält:
//! - data: TEXT_MATRIX und CURRENCIES
//! - builder: Sheet-Builder für Sprachversionen

pub mod builder;
pub mod data;

pub use builder::build_sheet;
pub use data::{CURRENCIES, TEXT_MATRIX};
