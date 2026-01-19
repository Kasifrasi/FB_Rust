use rust_xlsxwriter::{Format, Workbook, XlsxError};
use std::time::Instant;

fn main() -> Result<(), XlsxError> {
    println!("Testing unlocked cell performance...\n");
    
    // Test 1: 1000 rows x 30 cols
    {
        let mut workbook = Workbook::new();
        let worksheet = workbook.add_worksheet();
        let unlocked = Format::new().set_unlocked();
        
        let start = Instant::now();
        for row in 0..1000u32 {
            for col in 0..30u16 {
                worksheet.write_blank(row, col, &unlocked)?;
            }
        }
        println!("1000x30 (30,000 cells): {:?}", start.elapsed());
    }
    
    // Test 2: 5000 rows x 30 cols  
    {
        let mut workbook = Workbook::new();
        let worksheet = workbook.add_worksheet();
        let unlocked = Format::new().set_unlocked();
        
        let start = Instant::now();
        for row in 0..5000u32 {
            for col in 0..30u16 {
                worksheet.write_blank(row, col, &unlocked)?;
            }
        }
        println!("5000x30 (150,000 cells): {:?}", start.elapsed());
    }
    
    // Test 3: 10000 rows x 30 cols
    {
        let mut workbook = Workbook::new();
        let worksheet = workbook.add_worksheet();
        let unlocked = Format::new().set_unlocked();
        
        let start = Instant::now();
        for row in 0..10000u32 {
            for col in 0..30u16 {
                worksheet.write_blank(row, col, &unlocked)?;
            }
        }
        println!("10000x30 (300,000 cells): {:?}", start.elapsed());
    }

    println!("\nDone!");
    Ok(())
}
