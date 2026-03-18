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

/// Create a test file with metadata rows above the real header.
pub fn create_with_metadata(path: &Path) {
    use rust_xlsxwriter::*;
    let mut workbook = Workbook::new();
    let sheet = workbook.add_worksheet().set_name("Data").unwrap();

    // Metadata rows
    sheet.write_string(0, 0, "Quarterly Report").unwrap();
    sheet.write_string(1, 0, "Generated 2024-01-01").unwrap();

    // Real header at row 2
    sheet.write_string(2, 0, "Name").unwrap();
    sheet.write_string(2, 1, "Value").unwrap();

    // Data
    sheet.write_string(3, 0, "Alice").unwrap();
    sheet.write_number(3, 1, 100.0).unwrap();
    sheet.write_string(4, 0, "Bob").unwrap();
    sheet.write_number(4, 1, 200.0).unwrap();

    workbook.save(path.to_str().unwrap()).unwrap();
}

/// Create a test file with diverse data for filter testing.
/// Sheet "Data" with 6 rows: State, City, Amount, Year, Status
pub fn create_filterable(path: &Path) {
    let mut workbook = Workbook::new();
    let sheet = workbook.add_worksheet().set_name("Data").unwrap();

    // Headers
    sheet.write_string(0, 0, "State").unwrap();
    sheet.write_string(0, 1, "City").unwrap();
    sheet.write_string(0, 2, "Amount").unwrap();
    sheet.write_string(0, 3, "Year").unwrap();
    sheet.write_string(0, 4, "Status").unwrap();

    // Row 1: CA, Los Angeles, 1500, 2023, Active
    sheet.write_string(1, 0, "CA").unwrap();
    sheet.write_string(1, 1, "Los Angeles").unwrap();
    sheet.write_number(1, 2, 1500.0).unwrap();
    sheet.write_number(1, 3, 2023.0).unwrap();
    sheet.write_string(1, 4, "Active").unwrap();

    // Row 2: NY, New York, 2000, 2023, Active
    sheet.write_string(2, 0, "NY").unwrap();
    sheet.write_string(2, 1, "New York").unwrap();
    sheet.write_number(2, 2, 2000.0).unwrap();
    sheet.write_number(2, 3, 2023.0).unwrap();
    sheet.write_string(2, 4, "Active").unwrap();

    // Row 3: CA, San Francisco, 800, 2024, Draft
    sheet.write_string(3, 0, "CA").unwrap();
    sheet.write_string(3, 1, "San Francisco").unwrap();
    sheet.write_number(3, 2, 800.0).unwrap();
    sheet.write_number(3, 3, 2024.0).unwrap();
    sheet.write_string(3, 4, "Draft").unwrap();

    // Row 4: TX, Houston, 1200, 2024, Active
    sheet.write_string(4, 0, "TX").unwrap();
    sheet.write_string(4, 1, "Houston").unwrap();
    sheet.write_number(4, 2, 1200.0).unwrap();
    sheet.write_number(4, 3, 2024.0).unwrap();
    sheet.write_string(4, 4, "Active").unwrap();

    // Row 5: NY, Albany, 500, 2023, Draft
    sheet.write_string(5, 0, "NY").unwrap();
    sheet.write_string(5, 1, "Albany").unwrap();
    sheet.write_number(5, 2, 500.0).unwrap();
    sheet.write_number(5, 3, 2023.0).unwrap();
    sheet.write_string(5, 4, "Draft").unwrap();

    // Row 6: FL, Miami, 3000, 2024, Active
    sheet.write_string(6, 0, "FL").unwrap();
    sheet.write_string(6, 1, "Miami").unwrap();
    sheet.write_number(6, 2, 3000.0).unwrap();
    sheet.write_number(6, 3, 2024.0).unwrap();
    sheet.write_string(6, 4, "Active").unwrap();

    workbook.save(path).unwrap();
}

/// Create a pair of files for positional diff testing.
/// File A: Name/Score — Alice/90, Bob/80, Charlie/70
/// File B: Name/Score — Alice/90, Charlie/70, Dana/85
/// Expected: Bob removed, Dana added
pub fn create_diff_pair(path_a: &Path, path_b: &Path) {
    // File A
    {
        let mut wb = Workbook::new();
        let ws = wb.add_worksheet().set_name("Data").unwrap();
        ws.write_string(0, 0, "Name").unwrap();
        ws.write_string(0, 1, "Score").unwrap();
        ws.write_string(1, 0, "Alice").unwrap();
        ws.write_number(1, 1, 90.0).unwrap();
        ws.write_string(2, 0, "Bob").unwrap();
        ws.write_number(2, 1, 80.0).unwrap();
        ws.write_string(3, 0, "Charlie").unwrap();
        ws.write_number(3, 1, 70.0).unwrap();
        wb.save(path_a).unwrap();
    }

    // File B
    {
        let mut wb = Workbook::new();
        let ws = wb.add_worksheet().set_name("Data").unwrap();
        ws.write_string(0, 0, "Name").unwrap();
        ws.write_string(0, 1, "Score").unwrap();
        ws.write_string(1, 0, "Alice").unwrap();
        ws.write_number(1, 1, 90.0).unwrap();
        ws.write_string(2, 0, "Charlie").unwrap();
        ws.write_number(2, 1, 70.0).unwrap();
        ws.write_string(3, 0, "Dana").unwrap();
        ws.write_number(3, 1, 85.0).unwrap();
        wb.save(path_b).unwrap();
    }
}

/// Create a pair of files for key-based diff testing.
/// File A: ID/Name/Score — "1"/Alice/90, "2"/Bob/80, "3"/Charlie/70
/// File B: ID/Name/Score — "1"/Alice/95, "2"/Bob/80, "4"/Dana/85
/// Expected: ID=1 modified (Score 90→95), ID=3 removed, ID=4 added
pub fn create_diff_pair_with_keys(path_a: &Path, path_b: &Path) {
    // File A
    {
        let mut wb = Workbook::new();
        let ws = wb.add_worksheet().set_name("Data").unwrap();
        ws.write_string(0, 0, "ID").unwrap();
        ws.write_string(0, 1, "Name").unwrap();
        ws.write_string(0, 2, "Score").unwrap();
        ws.write_string(1, 0, "1").unwrap();
        ws.write_string(1, 1, "Alice").unwrap();
        ws.write_number(1, 2, 90.0).unwrap();
        ws.write_string(2, 0, "2").unwrap();
        ws.write_string(2, 1, "Bob").unwrap();
        ws.write_number(2, 2, 80.0).unwrap();
        ws.write_string(3, 0, "3").unwrap();
        ws.write_string(3, 1, "Charlie").unwrap();
        ws.write_number(3, 2, 70.0).unwrap();
        wb.save(path_a).unwrap();
    }

    // File B
    {
        let mut wb = Workbook::new();
        let ws = wb.add_worksheet().set_name("Data").unwrap();
        ws.write_string(0, 0, "ID").unwrap();
        ws.write_string(0, 1, "Name").unwrap();
        ws.write_string(0, 2, "Score").unwrap();
        ws.write_string(1, 0, "1").unwrap();
        ws.write_string(1, 1, "Alice").unwrap();
        ws.write_number(1, 2, 95.0).unwrap();
        ws.write_string(2, 0, "2").unwrap();
        ws.write_string(2, 1, "Bob").unwrap();
        ws.write_number(2, 2, 80.0).unwrap();
        ws.write_string(3, 0, "4").unwrap();
        ws.write_string(3, 1, "Dana").unwrap();
        ws.write_number(3, 2, 85.0).unwrap();
        wb.save(path_b).unwrap();
    }
}

/// Create a pair of files for tolerance testing.
/// File A: ID/Price — "1"/100.001, "2"/200.5
/// File B: ID/Price — "1"/100.002, "2"/200.6
/// Expected with tolerance 0.01: only ID=2 modified (diff=0.1 > 0.01)
pub fn create_diff_pair_with_floats(path_a: &Path, path_b: &Path) {
    // File A
    {
        let mut wb = Workbook::new();
        let ws = wb.add_worksheet().set_name("Data").unwrap();
        ws.write_string(0, 0, "ID").unwrap();
        ws.write_string(0, 1, "Price").unwrap();
        ws.write_string(1, 0, "1").unwrap();
        ws.write_number(1, 1, 100.001).unwrap();
        ws.write_string(2, 0, "2").unwrap();
        ws.write_number(2, 1, 200.5).unwrap();
        wb.save(path_a).unwrap();
    }

    // File B
    {
        let mut wb = Workbook::new();
        let ws = wb.add_worksheet().set_name("Data").unwrap();
        ws.write_string(0, 0, "ID").unwrap();
        ws.write_string(0, 1, "Price").unwrap();
        ws.write_string(1, 0, "1").unwrap();
        ws.write_number(1, 1, 100.002).unwrap();
        ws.write_string(2, 0, "2").unwrap();
        ws.write_number(2, 1, 200.6).unwrap();
        wb.save(path_b).unwrap();
    }
}
