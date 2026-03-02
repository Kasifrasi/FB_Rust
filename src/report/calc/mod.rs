//! IronCalc-basierte Formel-Engine
//!
//! Ersetzt die manuelle `CellRegistry` + `DynEvaluator` Architektur durch
//! eine IronCalc Spreadsheet-Engine. Formeln werden einmal definiert und
//! sowohl für Berechnung als auch Excel-Export verwendet.
//!
//! ## Module
//!
//! - `template`: `ModelTemplate` — Master-Modell mit statischen Formeln + Sprachversionen
//! - `bridge`: `CalcBridge` — Dynamische Formeln, Input-Population, Evaluation

mod bridge;
mod template;

pub(crate) use bridge::CalcBridge;
pub(crate) use template::ModelTemplate;
