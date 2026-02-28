//! Kern-Logik für die Formel-Evaluierung
//!
//! Dieses Modul enthält die interne Logik:
//! - `registry`: `CellRegistry`, `CellAddr` — zentrale Zellverwaltung mit topologischer Sortierung
//! - `definitions`: `build_registry()` — Registry-Builder und VLOOKUP-Evaluierung

mod definitions;
mod registry;

pub(crate) use definitions::build_registry;
pub(crate) use registry::{
    CellAddr, CellKind, CellRegistry, DynRegistry, EvalContext,
    FormulaCell, FormulaCellDeps, FormulaDeps, Inputs, RegistryError, Sheets, Statics,
};
