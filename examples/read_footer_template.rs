//! Liest den Footer-Bereich (B164:G184) aus der Template-Datei

use calamine::{open_workbook, Reader, Xlsx};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let path = "src/v2/Excel-Template/template.xlsx";
    let mut workbook: Xlsx<_> = open_workbook(path)?;

    // Erstes Sheet
    let sheet_names = workbook.sheet_names().to_vec();
    println!("Sheets: {:?}\n", sheet_names);

    // Formeln auslesen
    if let Ok(formula_range) = workbook.worksheet_formula(&sheet_names[0]) {
        println!("=== FORMELN im Footer-Bereich B164:G184 ===\n");

        for row in 163..184 {
            for col in 1..7 {
                // B=1 bis G=6
                let col_letter = match col {
                    1 => 'B',
                    2 => 'C',
                    3 => 'D',
                    4 => 'E',
                    5 => 'F',
                    6 => 'G',
                    _ => '?',
                };
                if let Some(formula) = formula_range.get_value((row as u32, col as u32)) {
                    if !formula.is_empty() {
                        println!("  {}{}: ={}", col_letter, row + 1, formula);
                    }
                }
            }
        }
    }

    // Werte auslesen
    if let Ok(range) = workbook.worksheet_range(&sheet_names[0]) {
        println!("\n=== WERTE im Footer-Bereich B164:G184 ===\n");

        for row in 163..184 {
            println!("--- Zeile {} (Excel {}) ---", row, row + 1);
            for col in 1..7 {
                // B=1 bis G=6
                let col_letter = match col {
                    1 => 'B',
                    2 => 'C',
                    3 => 'D',
                    4 => 'E',
                    5 => 'F',
                    6 => 'G',
                    _ => '?',
                };
                if let Some(cell) = range.get_value((row as u32, col as u32)) {
                    println!("  {}{}: {:?}", col_letter, row + 1, cell);
                }
            }
            println!();
        }
    }

    Ok(())
}
