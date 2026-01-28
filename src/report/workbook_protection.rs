//! Workbook Structure Protection
//!
//! rust_xlsxwriter unterstützt keinen Workbook-Strukturschutz direkt.
//! Dieses Modul bietet Post-Processing-Funktionen, um den Schutz
//! nach dem Speichern der XLSX-Datei hinzuzufügen.
//!
//! Der Workbook-Strukturschutz verhindert:
//! - Hinzufügen neuer Sheets
//! - Löschen von Sheets
//! - Umbenennen von Sheets
//! - Verschieben von Sheets
//! - Ausblenden/Einblenden von Sheets
//!
//! ## Verwendung
//!
//! ```ignore
//! use kmw_fb_rust::WorkbookProtection;
//!
//! // Nach workbook.save("report.xlsx")
//! WorkbookProtection::new()
//!     .with_password("secret")
//!     .lock_structure(true)
//!     .apply_to_file("report.xlsx")?;
//! ```

use std::fs::File;
use std::io::{Read, Write};
use std::path::Path;
use zip::write::SimpleFileOptions;
use zip::{ZipArchive, ZipWriter};

/// Workbook-Strukturschutz Konfiguration
#[derive(Debug, Clone, Default)]
pub struct WorkbookProtection {
    /// Passwort für den Schutz (optional, schwacher Schutz)
    pub password: Option<String>,
    /// Struktur sperren (Sheets hinzufügen/löschen/umbenennen verhindern)
    pub lock_structure: bool,
    /// Fenster sperren (Fenstergröße/-position verhindern)
    pub lock_windows: bool,
}

impl WorkbookProtection {
    /// Erstellt eine neue Workbook-Protection Konfiguration
    pub fn new() -> Self {
        Self {
            password: None,
            lock_structure: true, // Standard: Struktur sperren
            lock_windows: false,
        }
    }

    /// Setzt ein Passwort (schwacher Schutz - leicht zu umgehen)
    pub fn with_password(mut self, password: impl Into<String>) -> Self {
        self.password = Some(password.into());
        self
    }

    /// Setzt ob die Struktur gesperrt werden soll
    pub fn lock_structure(mut self, lock: bool) -> Self {
        self.lock_structure = lock;
        self
    }

    /// Setzt ob die Fenster gesperrt werden sollen
    pub fn lock_windows(mut self, lock: bool) -> Self {
        self.lock_windows = lock;
        self
    }

    /// Wendet den Workbook-Schutz auf eine existierende XLSX-Datei an
    ///
    /// Diese Funktion modifiziert die XLSX-Datei (ZIP-Archiv) und fügt
    /// das `<workbookProtection>` Element in `xl/workbook.xml` ein.
    pub fn apply_to_file<P: AsRef<Path>>(&self, path: P) -> Result<(), WorkbookProtectionError> {
        let path = path.as_ref();

        // 1. Originaldatei lesen
        let file = File::open(path).map_err(WorkbookProtectionError::IoError)?;
        let mut archive = ZipArchive::new(file).map_err(WorkbookProtectionError::ZipError)?;

        // 2. Temporäre Datei erstellen
        let temp_path = path.with_extension("xlsx.tmp");
        let temp_file = File::create(&temp_path).map_err(WorkbookProtectionError::IoError)?;
        let mut zip_writer = ZipWriter::new(temp_file);

        // 3. Alle Dateien kopieren, workbook.xml modifizieren
        for i in 0..archive.len() {
            let mut file = archive
                .by_index(i)
                .map_err(WorkbookProtectionError::ZipError)?;
            let name = file.name().to_string();

            let options = SimpleFileOptions::default()
                .compression_method(file.compression())
                .last_modified_time(
                    file.last_modified()
                        .unwrap_or(zip::DateTime::default_for_write()),
                );

            if name == "xl/workbook.xml" {
                // workbook.xml modifizieren
                let mut content = String::new();
                file.read_to_string(&mut content)
                    .map_err(WorkbookProtectionError::IoError)?;

                let modified = self.inject_protection(&content)?;

                zip_writer
                    .start_file(&name, options)
                    .map_err(WorkbookProtectionError::ZipError)?;
                zip_writer
                    .write_all(modified.as_bytes())
                    .map_err(WorkbookProtectionError::IoError)?;
            } else {
                // Andere Dateien unverändert kopieren
                zip_writer
                    .start_file(&name, options)
                    .map_err(WorkbookProtectionError::ZipError)?;

                let mut buffer = Vec::new();
                file.read_to_end(&mut buffer)
                    .map_err(WorkbookProtectionError::IoError)?;
                zip_writer
                    .write_all(&buffer)
                    .map_err(WorkbookProtectionError::IoError)?;
            }
        }

        // 4. ZIP abschließen
        zip_writer
            .finish()
            .map_err(WorkbookProtectionError::ZipError)?;

        // 5. Original durch modifizierte Datei ersetzen
        std::fs::rename(&temp_path, path).map_err(WorkbookProtectionError::IoError)?;

        Ok(())
    }

    /// Wendet den Workbook-Schutz auf einen Byte-Buffer an
    ///
    /// Nimmt die XLSX-Daten als Bytes und gibt die modifizierten Daten zurück.
    pub fn apply_to_buffer(&self, data: &[u8]) -> Result<Vec<u8>, WorkbookProtectionError> {
        use std::io::Cursor;

        let cursor = Cursor::new(data);
        let mut archive = ZipArchive::new(cursor).map_err(WorkbookProtectionError::ZipError)?;

        let mut output = Vec::new();
        {
            let cursor_out = Cursor::new(&mut output);
            let mut zip_writer = ZipWriter::new(cursor_out);

            for i in 0..archive.len() {
                let mut file = archive
                    .by_index(i)
                    .map_err(WorkbookProtectionError::ZipError)?;
                let name = file.name().to_string();

                let options = SimpleFileOptions::default()
                    .compression_method(file.compression())
                    .last_modified_time(
                        file.last_modified()
                            .unwrap_or(zip::DateTime::default_for_write()),
                    );

                if name == "xl/workbook.xml" {
                    let mut content = String::new();
                    file.read_to_string(&mut content)
                        .map_err(WorkbookProtectionError::IoError)?;

                    let modified = self.inject_protection(&content)?;

                    zip_writer
                        .start_file(&name, options)
                        .map_err(WorkbookProtectionError::ZipError)?;
                    zip_writer
                        .write_all(modified.as_bytes())
                        .map_err(WorkbookProtectionError::IoError)?;
                } else {
                    zip_writer
                        .start_file(&name, options)
                        .map_err(WorkbookProtectionError::ZipError)?;

                    let mut buffer = Vec::new();
                    file.read_to_end(&mut buffer)
                        .map_err(WorkbookProtectionError::IoError)?;
                    zip_writer
                        .write_all(&buffer)
                        .map_err(WorkbookProtectionError::IoError)?;
                }
            }

            zip_writer
                .finish()
                .map_err(WorkbookProtectionError::ZipError)?;
        }

        Ok(output)
    }

    /// Fügt das workbookProtection Element in den XML-String ein
    fn inject_protection(&self, xml: &str) -> Result<String, WorkbookProtectionError> {
        // Erstelle das workbookProtection Element
        let protection_element = self.build_protection_element();

        // Finde die Position zum Einfügen (nach <workbook ...> aber vor </workbook>)
        // Das Element sollte nach <fileVersion>, <workbookPr>, etc. kommen
        // aber vor <sheets>

        // Suche nach <sheets> als Einfügepunkt
        if let Some(sheets_pos) = xml.find("<sheets") {
            let mut result = String::with_capacity(xml.len() + protection_element.len() + 10);
            result.push_str(&xml[..sheets_pos]);
            result.push_str(&protection_element);
            result.push_str("\n  ");
            result.push_str(&xml[sheets_pos..]);
            Ok(result)
        } else {
            // Fallback: vor </workbook> einfügen
            if let Some(end_pos) = xml.find("</workbook>") {
                let mut result = String::with_capacity(xml.len() + protection_element.len() + 10);
                result.push_str(&xml[..end_pos]);
                result.push_str(&protection_element);
                result.push('\n');
                result.push_str(&xml[end_pos..]);
                Ok(result)
            } else {
                Err(WorkbookProtectionError::InvalidWorkbook(
                    "Could not find insertion point in workbook.xml".to_string(),
                ))
            }
        }
    }

    /// Erstellt das XML-Element für workbookProtection
    fn build_protection_element(&self) -> String {
        let mut attrs = Vec::new();

        if self.lock_structure {
            attrs.push("lockStructure=\"1\"".to_string());
        }

        if self.lock_windows {
            attrs.push("lockWindows=\"1\"".to_string());
        }

        if let Some(ref password) = self.password {
            let hash = hash_password(password);
            attrs.push(format!("workbookPassword=\"{:04X}\"", hash));
        }

        if attrs.is_empty() {
            String::new()
        } else {
            format!("<workbookProtection {}/>", attrs.join(" "))
        }
    }
}

/// Fehler beim Anwenden des Workbook-Schutzes
#[derive(Debug)]
pub enum WorkbookProtectionError {
    /// IO-Fehler
    IoError(std::io::Error),
    /// ZIP-Fehler
    ZipError(zip::result::ZipError),
    /// Ungültige Workbook-Struktur
    InvalidWorkbook(String),
}

impl std::fmt::Display for WorkbookProtectionError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            WorkbookProtectionError::IoError(e) => write!(f, "IO error: {}", e),
            WorkbookProtectionError::ZipError(e) => write!(f, "ZIP error: {}", e),
            WorkbookProtectionError::InvalidWorkbook(msg) => {
                write!(f, "Invalid workbook: {}", msg)
            }
        }
    }
}

impl std::error::Error for WorkbookProtectionError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            WorkbookProtectionError::IoError(e) => Some(e),
            WorkbookProtectionError::ZipError(e) => Some(e),
            WorkbookProtectionError::InvalidWorkbook(_) => None,
        }
    }
}

/// Hash-Funktion für Excel-Passwörter
///
/// Basiert auf dem Algorithmus in ECMA-376-4:2016, Office Open XML File Formats
/// — Transitional Migration Features, Additional attributes for workbookProtection
/// element (Part 1, §18.2.29).
fn hash_password(password: &str) -> u16 {
    let mut hash: u16 = 0;
    let length = password.len() as u16;

    if password.is_empty() {
        return 0;
    }

    for byte in password.as_bytes().iter().rev() {
        hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7FFF);
        hash ^= u16::from(*byte);
    }

    hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7FFF);
    hash ^= length;
    hash ^= 0xCE4B;

    hash
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_workbook_protection_builder() {
        let prot = WorkbookProtection::new()
            .with_password("test123")
            .lock_structure(true)
            .lock_windows(true);

        assert_eq!(prot.password, Some("test123".to_string()));
        assert!(prot.lock_structure);
        assert!(prot.lock_windows);
    }

    #[test]
    fn test_protection_element_no_password() {
        let prot = WorkbookProtection::new().lock_structure(true);

        let element = prot.build_protection_element();
        assert!(element.contains("lockStructure=\"1\""));
        assert!(!element.contains("workbookPassword"));
    }

    #[test]
    fn test_protection_element_with_password() {
        let prot = WorkbookProtection::new()
            .with_password("test")
            .lock_structure(true);

        let element = prot.build_protection_element();
        assert!(element.contains("lockStructure=\"1\""));
        assert!(element.contains("workbookPassword="));
    }

    #[test]
    fn test_hash_password() {
        // Bekannte Test-Werte
        assert_eq!(hash_password(""), 0);
        assert_ne!(hash_password("test"), 0);
        assert_ne!(hash_password("password"), 0);

        // Gleiche Passwörter sollten gleichen Hash haben
        assert_eq!(hash_password("abc"), hash_password("abc"));

        // Verschiedene Passwörter sollten verschiedene Hashes haben
        assert_ne!(hash_password("abc"), hash_password("def"));
    }

    #[test]
    fn test_inject_protection() {
        let prot = WorkbookProtection::new().lock_structure(true);

        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

        let result = prot.inject_protection(xml).unwrap();
        assert!(result.contains("<workbookProtection"));
        assert!(result.contains("lockStructure=\"1\""));
        // Protection sollte vor <sheets> eingefügt werden
        let prot_pos = result.find("<workbookProtection").unwrap();
        let sheets_pos = result.find("<sheets").unwrap();
        assert!(prot_pos < sheets_pos);
    }
}
