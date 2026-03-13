use rust_xlsxwriter::*;
use std::path::Path;

/// Single sheet, 5 rows of mixed types: string, float, int, bool
pub fn create_simple(path: &Path) {
    let mut wb = Workbook::new();
    let ws = wb.add_worksheet().set_name("Data").unwrap();

    // Headers
    ws.write_string(0, 0, "name").unwrap();
    ws.write_string(0, 1, "amount").unwrap();
    ws.write_string(0, 2, "count").unwrap();
    ws.write_string(0, 3, "active").unwrap();

    // Row 1
    ws.write_string(1, 0, "Alice").unwrap();
    ws.write_number(1, 1, 100.50).unwrap();
    ws.write_number(1, 2, 10.0).unwrap();
    ws.write_boolean(1, 3, true).unwrap();

    // Row 2
    ws.write_string(2, 0, "Bob").unwrap();
    ws.write_number(2, 1, 200.75).unwrap();
    ws.write_number(2, 2, 20.0).unwrap();
    ws.write_boolean(2, 3, false).unwrap();

    // Row 3
    ws.write_string(3, 0, "Charlie").unwrap();
    ws.write_number(3, 1, 300.00).unwrap();
    ws.write_number(3, 2, 30.0).unwrap();
    ws.write_boolean(3, 3, true).unwrap();

    // Row 4
    ws.write_string(4, 0, "Diana").unwrap();
    ws.write_number(4, 1, 400.25).unwrap();
    ws.write_number(4, 2, 40.0).unwrap();
    ws.write_boolean(4, 3, false).unwrap();

    // Row 5
    ws.write_string(5, 0, "Eve").unwrap();
    ws.write_number(5, 1, 500.00).unwrap();
    ws.write_number(5, 2, 50.0).unwrap();
    ws.write_boolean(5, 3, true).unwrap();

    wb.save(path).unwrap();
}

/// 3 sheets: Revenue (4 rows), Expenses (3 rows), Summary (2 rows)
pub fn create_multi_sheet(path: &Path) {
    let mut wb = Workbook::new();

    let ws1 = wb.add_worksheet().set_name("Revenue").unwrap();
    ws1.write_string(0, 0, "region").unwrap();
    ws1.write_string(0, 1, "amount").unwrap();
    for i in 1..=4u32 {
        ws1.write_string(i, 0, &format!("Region {i}")).unwrap();
        ws1.write_number(i, 1, i as f64 * 1000.0).unwrap();
    }

    let ws2 = wb.add_worksheet().set_name("Expenses").unwrap();
    ws2.write_string(0, 0, "category").unwrap();
    ws2.write_string(0, 1, "amount").unwrap();
    for i in 1..=3u32 {
        ws2.write_string(i, 0, &format!("Category {i}")).unwrap();
        ws2.write_number(i, 1, i as f64 * 500.0).unwrap();
    }

    let ws3 = wb.add_worksheet().set_name("Summary").unwrap();
    ws3.write_string(0, 0, "metric").unwrap();
    ws3.write_string(0, 1, "value").unwrap();
    ws3.write_string(1, 0, "Total Revenue").unwrap();
    ws3.write_number(1, 1, 10000.0).unwrap();
    ws3.write_string(2, 0, "Total Expenses").unwrap();
    ws3.write_number(2, 1, 3000.0).unwrap();

    wb.save(path).unwrap();
}

/// Single sheet with 80 rows (to test head/tail adaptive behavior)
pub fn create_many_rows(path: &Path) {
    let mut wb = Workbook::new();
    let ws = wb.add_worksheet().set_name("Data").unwrap();

    ws.write_string(0, 0, "id").unwrap();
    ws.write_string(0, 1, "value").unwrap();

    for i in 1..=80u32 {
        ws.write_number(i, 0, i as f64).unwrap();
        ws.write_number(i, 1, i as f64 * 1.5).unwrap();
    }

    wb.save(path).unwrap();
}

/// Single sheet with header row but no data rows
pub fn create_empty_data(path: &Path) {
    let mut wb = Workbook::new();
    let ws = wb.add_worksheet().set_name("Empty").unwrap();
    ws.write_string(0, 0, "col_a").unwrap();
    ws.write_string(0, 1, "col_b").unwrap();
    wb.save(path).unwrap();
}

/// Completely empty sheet
pub fn create_empty_sheet(path: &Path) {
    let mut wb = Workbook::new();
    wb.add_worksheet().set_name("Blank").unwrap();
    wb.save(path).unwrap();
}
