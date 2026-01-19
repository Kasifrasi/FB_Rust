//! Test: set_row_format() mit unlocked für 10000 Zeilen

use rust_xlsxwriter::{Format, Workbook};
use std::time::Instant;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("╔══════════════════════════════════════════════════════════════╗");
    println!("║   TEST: set_row_format() mit unlocked                        ║");
    println!("╚══════════════════════════════════════════════════════════════╝\n");

    // Test 1: 10000 Zeilen mit set_row_format()
    println!("📊 Test 1: set_row_format() für 10000 Zeilen");
    let start = Instant::now();
    
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    
    let unlocked = Format::new()
        .set_font_name("Arial")
        .set_font_size(10.0)
        .set_unlocked();
    
    for row in 0..10000u32 {
        worksheet.set_row_format(row, &unlocked)?;
    }
    
    // Schreibe einige Testdaten
    worksheet.write_string(0, 0, "Test A1")?;
    worksheet.write_string(5, 2, "Test C6")?;
    worksheet.write_number(100, 5, 123.45)?;
    
    // Protect
    worksheet.protect();
    
    workbook.save("/tmp/test_row_format.xlsx")?;
    
    let duration1 = start.elapsed();
    println!("   ✅ Dauer: {:?}\n", duration1);

    // Test 2: Vergleich mit write_blank() für 10000x30
    println!("📊 Test 2: write_blank() für 10000x30 Zellen (Vergleich)");
    let start = Instant::now();
    
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    
    let unlocked = Format::new()
        .set_font_name("Arial")
        .set_font_size(10.0)
        .set_unlocked();
    
    for row in 0..10000u32 {
        for col in 0..30u16 {
            worksheet.write_blank(row, col, &unlocked)?;
        }
    }
    
    worksheet.write_string(0, 0, "Test A1")?;
    worksheet.write_string(5, 2, "Test C6")?;
    worksheet.write_number(100, 5, 123.45)?;
    
    worksheet.protect();
    
    workbook.save("/tmp/test_write_blank.xlsx")?;
    
    let duration2 = start.elapsed();
    println!("   ✅ Dauer: {:?}\n", duration2);

    // Test 3: unprotect_range() (aktueller Ansatz)
    println!("📊 Test 3: unprotect_range() (Vergleich)");
    let start = Instant::now();
    
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    
    worksheet.write_string(0, 0, "Test A1")?;
    worksheet.write_string(5, 2, "Test C6")?;
    worksheet.write_number(100, 5, 123.45)?;
    
    worksheet.protect();
    worksheet.unprotect_range(0, 0, 9999, 9999)?;
    
    workbook.save("/tmp/test_unprotect_range.xlsx")?;
    
    let duration3 = start.elapsed();
    println!("   ✅ Dauer: {:?}\n", duration3);

    // Ergebnisse
    println!("═══════════════════════════════════════════════════════════════");
    println!("                    VERGLEICH");
    println!("═══════════════════════════════════════════════════════════════\n");
    
    println!("  set_row_format() 10000 Zeilen : {:?}", duration1);
    println!("  write_blank() 10000x30 Zellen : {:?}", duration2);
    println!("  unprotect_range() 10000x10000 : {:?}", duration3);
    
    let speedup_vs_blank = duration2.as_secs_f64() / duration1.as_secs_f64();
    println!("\n  → set_row_format() ist {:.1}x schneller als write_blank()", speedup_vs_blank);
    
    let speedup_vs_unprotect = duration3.as_secs_f64() / duration1.as_secs_f64();
    if speedup_vs_unprotect > 1.0 {
        println!("  → set_row_format() ist {:.1}x schneller als unprotect_range()", speedup_vs_unprotect);
    } else {
        println!("  → unprotect_range() ist {:.1}x schneller als set_row_format()", 1.0/speedup_vs_unprotect);
    }

    println!("\n🧹 Dateien erstellt:");
    println!("   /tmp/test_row_format.xlsx");
    println!("   /tmp/test_write_blank.xlsx");
    println!("   /tmp/test_unprotect_range.xlsx");
    println!("\n💡 Bitte prüfe in Excel, ob die Zellen wirklich 'unlocked' sind!");

    Ok(())
}
