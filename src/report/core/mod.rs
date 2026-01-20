//! Kern-Logik für die Formel-Evaluierung
//!
//! Dieses Modul enthält die interne Logik:
//! - `registry`: CellRegistry mit topologischer Sortierung
//! - `cells`: CellAddress für Zelladressen
//! - `definitions`: Registry-Builder und VLOOKUP-Funktionen
//! - `formulas`: Formel-Evaluierung und Caching
//! - `dynamic`: Dynamische Sections (für Body)

mod cells;
mod definitions;
mod dynamic;
mod formulas;
mod registry;

pub use cells::CellAddress;
pub use definitions::{build_registry, lookup_text, lookup_text_string};
pub use dynamic::{DynamicRegistry, DynamicSectionDef};
pub use formulas::{
    evaluate_formula, FormulaCache, FormulaDefinition, FormulaType, HEADER_FORMULAS,
};
pub use registry::{
    CellAddr, CellEvaluator, CellKind, CellRegistry, EvalContext, FormulaCell, FormulaCellDeps,
    FormulaDeps, Inputs, RegistryError, SheetRef, Sheets, StaticVal, Statics,
};
