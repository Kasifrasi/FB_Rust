//! Test: Formeln mit absichtlich FALSCHEN Cache-Werten
//!
//! Generiert eine Excel-Datei, bei der die Formel-Ergebnisse bewusst falsch
//! gecacht sind. Damit lässt sich testen, ob Excel beim Öffnen automatisch
//! die Formeln neu berechnet und die korrekten Werte anzeigt.
//!
//! Erwartetes Verhalten:
//! - Beim Öffnen in Excel: Formeln werden sofort neu berechnet → korrekte Werte
//! - Beim Öffnen in LibreOffice: Ebenfalls Neuberechnung (Standard)
//! - Ohne Neuberechnung: Die falschen Cache-Werte werden angezeigt

use rust_xlsxwriter::{Format, Formula, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    println!("Erstelle Excel mit falschen Formel-Caches...\n");

    let mut workbook = Workbook::new();
    let ws = workbook.add_worksheet();
    ws.set_name("Formel-Cache-Test")?;

    let bold = Format::new().set_bold();
    let number = Format::new().set_num_format("#,##0.00");

    // === Header ===
    ws.write_string_with_format(0, 0, "Beschreibung", &bold)?;
    ws.write_string_with_format(0, 1, "Wert A", &bold)?;
    ws.write_string_with_format(0, 2, "Wert B", &bold)?;
    ws.write_string_with_format(0, 3, "Formel", &bold)?;
    ws.write_string_with_format(0, 4, "Cache (FALSCH)", &bold)?;
    ws.write_string_with_format(0, 5, "Erwartet (RICHTIG)", &bold)?;

    // === Test 1: Einfache Addition ===
    ws.write_string(1, 0, "A + B")?;
    ws.write_number(1, 1, 100.0)?;
    ws.write_number(1, 2, 200.0)?;
    // Korrekt wäre 300, Cache sagt 999
    let f1 = Formula::new("=B2+C2").set_result("999");
    ws.write_formula_with_format(1, 3, f1, &number)?;
    ws.write_string(1, 4, "999 (FALSCH)")?;
    ws.write_string(1, 5, "300 (RICHTIG)")?;

    // === Test 2: SUM ===
    ws.write_string(2, 0, "SUM(A, B)")?;
    ws.write_number(2, 1, 1_000.0)?;
    ws.write_number(2, 2, 2_000.0)?;
    // Korrekt wäre 3000, Cache sagt 42
    let f2 = Formula::new("=SUM(B3:C3)").set_result("42");
    ws.write_formula_with_format(2, 3, f2, &number)?;
    ws.write_string(2, 4, "42 (FALSCH)")?;
    ws.write_string(2, 5, "3.000 (RICHTIG)")?;

    // === Test 3: Multiplikation ===
    ws.write_string(3, 0, "A * B")?;
    ws.write_number(3, 1, 50.0)?;
    ws.write_number(3, 2, 3.0)?;
    // Korrekt wäre 150, Cache sagt -1
    let f3 = Formula::new("=B4*C4").set_result("-1");
    ws.write_formula_with_format(3, 3, f3, &number)?;
    ws.write_string(3, 4, "-1 (FALSCH)")?;
    ws.write_string(3, 5, "150 (RICHTIG)")?;

    // === Test 4: Prozent ===
    ws.write_string(4, 0, "A / B (Anteil)")?;
    ws.write_number(4, 1, 75.0)?;
    ws.write_number(4, 2, 300.0)?;
    // Korrekt wäre 0.25 (25%), Cache sagt 0.99
    let pct = Format::new().set_num_format("0%");
    let f4 = Formula::new("=B5/C5").set_result("0.99");
    ws.write_formula_with_format(4, 3, f4, &pct)?;
    ws.write_string(4, 4, "99% (FALSCH)")?;
    ws.write_string(4, 5, "25% (RICHTIG)")?;

    // === Test 5: IF-Bedingung ===
    ws.write_string(5, 0, "IF(A>100, \"Ja\", \"Nein\")")?;
    ws.write_number(5, 1, 50.0)?;
    // Korrekt wäre "Nein" (50 < 100), Cache sagt "Ja"
    let f5 = Formula::new("=IF(B6>100,\"Ja\",\"Nein\")").set_result("Ja");
    ws.write_formula(5, 3, f5)?;
    ws.write_string(5, 4, "Ja (FALSCH)")?;
    ws.write_string(5, 5, "Nein (RICHTIG)")?;

    // === Test 6: Leerer Cache vs. Zahl ===
    ws.write_string(6, 0, "A + 1 (Cache leer)")?;
    ws.write_number(6, 1, 10.0)?;
    // Korrekt wäre 11, Cache ist leer ("")
    let f6 = Formula::new("=B7+1").set_result("");
    ws.write_formula_with_format(6, 3, f6, &number)?;
    ws.write_string(6, 4, "(leer) (FALSCH)")?;
    ws.write_string(6, 5, "11 (RICHTIG)")?;

    // === Spaltenbreiten ===
    ws.set_column_width(0, 30.0)?;
    ws.set_column_width(1, 12.0)?;
    ws.set_column_width(2, 12.0)?;
    ws.set_column_width(3, 15.0)?;
    ws.set_column_width(4, 20.0)?;
    ws.set_column_width(5, 22.0)?;

    let path = "examples/output/test_wrong_cache.xlsx";
    workbook.save(path)?;

    println!("Datei gespeichert: {path}");
    println!();
    println!("Spalte D enthält Formeln mit FALSCHEN Cache-Werten.");
    println!("Wenn Excel beim Öffnen automatisch neu berechnet,");
    println!("sollten die Werte in Spalte D mit Spalte F übereinstimmen.");

    Ok(())
}
