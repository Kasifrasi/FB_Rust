//! Semantische Fill-Farben mit Bedeutung
//!
//! Diese Farben werden als VARIABLEN verwendet, damit eine zentrale
//! Änderung alle Zellen mit dieser Bedeutung ändert (gewollter Welleneffekt).

use rust_xlsxwriter::Color;

/// Semantische Fill-Farben mit fester Bedeutung
///
/// # Verwendung
///
/// ```ignore
/// let fills = FillColors::new();
///
/// // Fill via Variable in Zell-Definition
/// ws.write_string(0, 4, "Value",
///     &Format::new()
///         .set_font_name("Arial")
///         .set_font_size(10.0)
///         .set_background_color(fills.input)  // ✅ Variable!
///         .set_unlocked()
/// )?;
/// ```
///
/// # Gewollter Welleneffekt
///
/// Alle Input-Felder von Gelb → Grün ändern:
/// ```ignore
/// let mut fills = FillColors::new();
/// fills.input = Color::RGB(0x00FF00);  // Eine Zeile - alle Input-Felder ändern sich!
/// ```
#[derive(Debug, Clone)]
pub struct FillColors {
    /// GELB: Input-Felder (Benutzer MUSS ausfüllen)
    pub input: Color,

    /// GRAU HELL: Berechnete Werte (Read-only Anzeige)
    pub value: Color,

    /// GRAU DUNKEL: Summary-Zeilen (Total, Summen)
    pub summary: Color,

    /// ORANGE: Highlight/Special (z.B. Projekttitel)
    pub highlight: Color,
}

impl FillColors {
    /// Erstellt Standard-Fill-Farben
    pub fn new() -> Self {
        Self {
            input: Color::RGB(0xFFFAE5),     // Gelb
            value: Color::RGB(0xF2F2F2),     // Grau hell
            summary: Color::RGB(0xD9D9D9),   // Grau dunkel
            highlight: Color::RGB(0xFFCC00), // Orange
        }
    }

    /// Gibt alle Farben mit ihrer Bedeutung zurück (für Dokumentation/Debugging)
    pub fn describe(&self) -> Vec<(&'static str, Color)> {
        vec![
            ("Input (Benutzer füllt aus)", self.input),
            ("Value (berechnete Anzeige)", self.value),
            ("Summary (Total-Zeilen)", self.summary),
            ("Highlight (Special)", self.highlight),
        ]
    }
}

impl Default for FillColors {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_fill_colors_default() {
        let fills = FillColors::new();
        assert_eq!(fills.input, Color::RGB(0xFFFAE5));
        assert_eq!(fills.value, Color::RGB(0xF2F2F2));
        assert_eq!(fills.summary, Color::RGB(0xD9D9D9));
        assert_eq!(fills.highlight, Color::RGB(0xFFCC00));
    }

    #[test]
    fn test_fill_colors_describe() {
        let fills = FillColors::new();
        let descriptions = fills.describe();

        assert_eq!(descriptions.len(), 4);
        assert_eq!(descriptions[0].0, "Input (Benutzer füllt aus)");
        assert_eq!(descriptions[1].0, "Value (berechnete Anzeige)");
        assert_eq!(descriptions[2].0, "Summary (Total-Zeilen)");
        assert_eq!(descriptions[3].0, "Highlight (Special)");
    }

    #[test]
    fn test_fill_colors_mutability() {
        let mut fills = FillColors::new();

        // Ändere Input-Farbe von Gelb → Grün
        fills.input = Color::RGB(0x00FF00);

        assert_eq!(fills.input, Color::RGB(0x00FF00));
        // Andere Farben bleiben unverändert
        assert_eq!(fills.value, Color::RGB(0xF2F2F2));
    }

    #[test]
    fn test_default_trait() {
        let fills: FillColors = Default::default();
        assert_eq!(fills.input, Color::RGB(0xFFFAE5));
    }
}
