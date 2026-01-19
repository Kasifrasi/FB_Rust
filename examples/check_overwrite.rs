use rust_xlsxwriter::{Color, Format, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    // Create some formats
    let red_bg = Format::new().set_background_color(Color::Red);
    let blue_bg = Format::new().set_background_color(Color::Blue);

    // Test 1: Write Format, then Value
    // Expectation: usually the second write completely replaces the first (erasing the format)
    // unless rust_xlsxwriter has specific behavior to merge.
    worksheet.write_blank(0, 0, &red_bg)?;
    worksheet.write_string(0, 0, "Text only (was red)")?;
    worksheet.write_string(
        0,
        1,
        "<-- Cell A1: Created red blank, then wrote string without format.",
    )?;

    // Test 2: Write Value, then Format (as blank)
    // Expectation: Value is lost, replaced by blank with format.
    worksheet.write_string(1, 0, "Value (will be blue)")?;
    worksheet.write_blank(1, 0, &blue_bg)?;
    worksheet.write_string(1, 1, "<-- Cell A2: Wrote string, then wrote blue blank.")?;

    // Test 3: Write with Format, then Write Formula without format
    worksheet.write_string_with_format(2, 0, "Formatted", &red_bg)?;
    worksheet.write_formula(2, 0, "=\"Formula Only\"")?;
    worksheet.write_string(
        2,
        1,
        "<-- Cell A3: Wrote red string, then wrote formula without format.",
    )?;

    workbook.save("check_overwrite.xlsx")?;

    println!("Created check_overwrite.xlsx");
    Ok(())
}
