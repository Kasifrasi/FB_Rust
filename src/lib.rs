//! KMW-FB Rust - Financial Report Generator
//!
//! High-performance Excel financial report generator with dynamic cost positions,
//! formulas, formatting, and multi-language support.

pub mod common;
pub mod lang;
pub mod report;

#[cfg(test)]
mod tests;

// Public API re-exports
pub use report::{
    api::ApiKey,
    body::BodyConfig,
    formats::ReportStyles,
    values::ReportValues,
    writer::{write_report, write_report_with_body},
};
