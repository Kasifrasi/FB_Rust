//! Zell-Adressen für den Finanzbericht
//!
//! HINWEIS: API-Zellen werden jetzt in `api.rs` definiert.
//! Diese Datei enthält nur noch die generische `CellAddress` Struktur.

use std::fmt;

/// Excel-Zelladresse (0-basiert)
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct CellAddress {
    pub row: u32,
    pub col: u16,
}

impl CellAddress {
    pub const fn new(row: u32, col: u16) -> Self {
        Self { row, col }
    }

    /// Konvertiert zu Excel-Notation (z.B. "E2")
    pub fn to_excel_notation(&self) -> String {
        let col_letter = Self::col_to_letter(self.col);
        format!("{}{}", col_letter, self.row + 1)
    }

    fn col_to_letter(col: u16) -> String {
        let mut result = String::new();
        let mut c = col as u32;
        loop {
            result.insert(0, (b'A' + (c % 26) as u8) as char);
            if c < 26 {
                break;
            }
            c = c / 26 - 1;
        }
        result
    }
}

impl fmt::Display for CellAddress {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{}", self.to_excel_notation())
    }
}

// ============================================================================
// Tests
// ============================================================================

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_cell_address_to_excel() {
        assert_eq!(CellAddress::new(0, 0).to_excel_notation(), "A1");
        assert_eq!(CellAddress::new(1, 4).to_excel_notation(), "E2");
        assert_eq!(CellAddress::new(13, 11).to_excel_notation(), "L14");
        assert_eq!(CellAddress::new(0, 25).to_excel_notation(), "Z1");
        assert_eq!(CellAddress::new(0, 26).to_excel_notation(), "AA1");
    }
}
