use rust_xlsxwriter::{Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    // Set zoom to 85%
    worksheet.set_zoom(85);

    worksheet.write_string(0, 0, "Zoom should be 85%")?;

    workbook.save("test_zoom.xlsx")?;

    println!("Created test_zoom.xlsx with zoom set to 85%");
    Ok(())
}
