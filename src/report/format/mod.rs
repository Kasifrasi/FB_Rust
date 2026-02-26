//! Formatierung und Schutz
//!
//! Dieses Modul enthält:
//! - `styles`: ReportStyles, FormatMatrix, SectionStyles
//! - `protection`: SheetProtection, Validation, HiddenRanges

mod protection;
mod styles;

pub use protection::{
    FieldValidation, HiddenRange, HiddenRanges, NumericRule, ReportOptions, RowGroup, RowGrouping,
    SheetProtection, ValidationError, ValidationErrorStyle, ValidationRule, ValidationRuleType,
    ValidationTarget, WorkbookProtection,
};
pub use styles::{
    build_format_matrix, extend_format_matrix_with_body, extend_format_matrix_with_footer,
    BodyStyles, FormatMatrix, ReportStyles, SectionStyles,
};
