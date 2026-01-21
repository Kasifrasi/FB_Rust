//! Formatierung und Schutz
//!
//! Dieses Modul enthält:
//! - `fills`: FillColors - Semantische Farben mit Bedeutung
//! - `borders`: BorderManager, BorderSpec, MergedCellRegistry - Border-System mit Merge-Awareness
//! - `styles`: ReportStyles, FormatMatrix, SectionStyles
//! - `protection`: SheetProtection, Validation, HiddenRanges

mod borders;
mod fills;
mod protection;
mod styles;

pub use borders::{BorderError, BorderManager, BorderSpec, MergedCellRegistry};
pub use fills::FillColors;
pub use protection::{
    FieldValidation, HiddenRange, HiddenRanges, NumericRule, ReportOptions, SheetProtection,
    ValidationError, ValidationErrorStyle, ValidationRule, ValidationRuleType, ValidationTarget,
};
pub use styles::{
    build_format_matrix, extend_format_matrix_with_body, BodyStyles, FormatMatrix, ReportStyles,
    SectionStyles,
};
