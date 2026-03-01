//! Language Module - Sprachversionen und Übersetzungen
//!
//! Enthält:
//! - data: TEXT_MATRIX
//! - builder: Sheet-Builder für Sprachversionen

pub mod builder;
pub mod config;
pub mod data;

pub use builder::{build_sheet, build_sheet_with_visibility};
pub use config::{LanguageConfig, LANG_CONFIG, LANG_SUFFIXES};
pub use data::TEXT_MATRIX;
