//! API-Typen für den Finanzbericht
//!
//! Dieses Modul enthält alle öffentlichen Typen für die Nutzung der API:
//! - `keys`: ApiKey und PositionField für Zell-Referenzierung
//! - `types`: Typsichere Typen (Language, Currency, Category, ReportDate)
//! - `values`: ReportValues für Eingabewerte

mod keys;
mod types;
mod values;

pub use keys::{register_all_api_cells, ApiKey, FooterField, PositionField};
pub use types::{Category, Currency, DateError, Language, ReportDate};
pub use values::{CellValue, ReportValues};
