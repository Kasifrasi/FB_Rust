//! Crate-level error types for report generation.

use crate::workbook_protection::ProtectionError;
use rust_xlsxwriter::XlsxError;
use std::fmt;

/// Top-Level-Fehler für die Report-Erstellung.
///
/// Wird von [`ReportConfig::write_to`](crate::config::ReportConfig::write_to),
/// [`ReportConfig::write_to_precomputed`](crate::config::ReportConfig::write_to_precomputed)
/// und den internen Engine-Funktionen zurückgegeben.
#[derive(Debug)]
pub enum ReportError {
    /// Fehler beim Excel-Schreiben (rust_xlsxwriter)
    Xlsx(XlsxError),
    /// Fehler bei Workbook-Protection (ZIP/XML-Manipulation)
    Protection(ProtectionError),
    /// I/O-Fehler (Temp-Datei, Persistierung etc.)
    Io(std::io::Error),
    /// Pfad enthält ungültige UTF-8-Zeichen
    InvalidPath(String),
}

impl fmt::Display for ReportError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Xlsx(e) => write!(f, "Excel error: {e}"),
            Self::Protection(e) => write!(f, "Workbook protection error: {e}"),
            Self::Io(e) => write!(f, "I/O error: {e}"),
            Self::InvalidPath(p) => write!(f, "Path contains non-UTF-8 characters: {p}"),
        }
    }
}

impl std::error::Error for ReportError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Xlsx(e) => Some(e),
            Self::Protection(e) => Some(e),
            Self::Io(e) => Some(e),
            Self::InvalidPath(_) => None,
        }
    }
}

impl From<XlsxError> for ReportError {
    fn from(e: XlsxError) -> Self {
        Self::Xlsx(e)
    }
}

impl From<ProtectionError> for ReportError {
    fn from(e: ProtectionError) -> Self {
        Self::Protection(e)
    }
}

impl From<std::io::Error> for ReportError {
    fn from(e: std::io::Error) -> Self {
        Self::Io(e)
    }
}

impl From<tempfile::PersistError> for ReportError {
    fn from(e: tempfile::PersistError) -> Self {
        Self::Io(e.error)
    }
}
