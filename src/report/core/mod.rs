//! Kern-Logik für die Formel-Evaluierung
//!
//! Dieses Modul enthält die interne Logik:
//! - `registry`: `CellRegistry`, `CellAddr` — zentrale Zellverwaltung mit topologischer Sortierung
//! - `definitions`: `build_registry()` — Registry-Builder und VLOOKUP-Evaluierung

mod definitions;
mod registry;

pub use definitions::{build_registry, lookup_text, lookup_text_string};
pub use registry::{
    CellAddr, CellEvaluator, CellKind, CellRegistry, DynEvaluator, DynRegistry, EvalContext,
    FormulaCell, FormulaCellDeps, FormulaDeps, Inputs, RegistryError, SheetRef, Sheets, StaticVal,
    Statics,
};
