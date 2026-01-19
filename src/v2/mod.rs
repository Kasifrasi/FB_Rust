pub mod common;
pub mod lang;
pub mod report;
mod tests;

// Legacy re-exports (für Abwärtskompatibilität während Übergang)
#[allow(non_snake_case)]
pub mod Finanzbericht {
    // Modules
    pub mod header {
        pub use crate::v2::report::writer::write_report as write_header;
    }
    pub mod sheet_setup {
        pub use crate::v2::report::layout::setup_sheet as sheet_setup;
    }
    pub mod styles {
        pub use crate::v2::report::formats::ReportStyles;
    }
    pub mod values {
        pub use crate::v2::report::values::ReportValues;
    }
}

#[allow(non_snake_case)]
pub mod Sprachversion {
    pub mod builder {
        pub use crate::v2::lang::builder::build_sheet;
    }
}
