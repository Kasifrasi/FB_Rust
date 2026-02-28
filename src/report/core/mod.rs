//! Kern-Logik für die Formel-Evaluierung
//!
//! Dieses Modul enthält die interne Logik:
//! - `registry`: `CellRegistry`, `CellAddr` — zentrale Zellverwaltung mit topologischer Sortierung
//! - `definitions`: `build_registry()` — Registry-Builder und VLOOKUP-Evaluierung
//! - `formulas`: `FormulaDefinition`, `FormulaCache` — Formeltypen und Cache
//! - `dynamic`: `DynamicRegistry` — Infrastruktur für variable Sektionen

mod definitions;
mod dynamic;
mod formulas;
mod registry;

pub use definitions::{build_registry, lookup_text, lookup_text_string};
pub use dynamic::{DynamicRegistry, DynamicSectionDef};
pub use formulas::{
    evaluate_formula, FormulaCache, FormulaDefinition, FormulaType, HEADER_FORMULAS,
};
pub use registry::{
    CellAddr, CellEvaluator, CellKind, CellRegistry, EvalContext, FormulaCell, FormulaCellDeps,
    FormulaDeps, Inputs, RegistryError, SheetRef, Sheets, StaticVal, Statics,
};
