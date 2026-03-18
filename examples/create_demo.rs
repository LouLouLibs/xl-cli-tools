use rust_xlsxwriter::*;
use std::path::Path;

fn main() {
    let dir = Path::new("demo");

    // --- sales.xlsx: single sheet, 12 rows ---
    {
        let mut wb = Workbook::new();
        let ws = wb.add_worksheet();
        ws.set_name("Q1 Sales").unwrap();

        ws.write_string(0, 0, "Date").unwrap();
        ws.write_string(0, 1, "Region").unwrap();
        ws.write_string(0, 2, "Product").unwrap();
        ws.write_string(0, 3, "Revenue").unwrap();
        ws.write_string(0, 4, "Units").unwrap();

        let data = [
            ("2024-01-05", "East",  "Widget A", 12500.0, 250.0),
            ("2024-01-12", "West",  "Widget B",  8700.0, 145.0),
            ("2024-01-19", "East",  "Widget A", 15300.0, 306.0),
            ("2024-01-26", "North", "Widget C",  4200.0,  84.0),
            ("2024-02-02", "West",  "Widget A", 11800.0, 236.0),
            ("2024-02-09", "East",  "Widget B",  9100.0, 152.0),
            ("2024-02-16", "South", "Widget C",  6700.0, 134.0),
            ("2024-02-23", "North", "Widget A", 13400.0, 268.0),
            ("2024-03-01", "West",  "Widget B",  7900.0, 132.0),
            ("2024-03-08", "East",  "Widget C",  5500.0, 110.0),
            ("2024-03-15", "South", "Widget A", 14200.0, 284.0),
            ("2024-03-22", "North", "Widget B", 10600.0, 177.0),
        ];

        for (i, (date, region, product, revenue, units)) in data.iter().enumerate() {
            let row = (i + 1) as u32;
            ws.write_string(row, 0, *date).unwrap();
            ws.write_string(row, 1, *region).unwrap();
            ws.write_string(row, 2, *product).unwrap();
            ws.write_number(row, 3, *revenue).unwrap();
            ws.write_number(row, 4, *units).unwrap();
        }

        wb.save(dir.join("sales.xlsx")).unwrap();
        println!("Created demo/sales.xlsx");
    }

    // --- budget.xlsx: 3 sheets ---
    {
        let mut wb = Workbook::new();

        let ws1 = wb.add_worksheet();
        ws1.set_name("Revenue").unwrap();
        ws1.write_string(0, 0, "Quarter").unwrap();
        ws1.write_string(0, 1, "Amount").unwrap();
        ws1.write_string(0, 2, "Target").unwrap();
        for (i, (q, amt, tgt)) in [
            ("Q1", 145000.0, 140000.0),
            ("Q2", 162000.0, 155000.0),
            ("Q3", 138000.0, 150000.0),
            ("Q4", 171000.0, 165000.0),
        ].iter().enumerate() {
            let row = (i + 1) as u32;
            ws1.write_string(row, 0, *q).unwrap();
            ws1.write_number(row, 1, *amt).unwrap();
            ws1.write_number(row, 2, *tgt).unwrap();
        }

        let ws2 = wb.add_worksheet();
        ws2.set_name("Expenses").unwrap();
        ws2.write_string(0, 0, "Category").unwrap();
        ws2.write_string(0, 1, "Q1").unwrap();
        ws2.write_string(0, 2, "Q2").unwrap();
        for (i, (cat, q1, q2)) in [
            ("Payroll",   85000.0, 87000.0),
            ("Marketing", 12000.0, 15000.0),
            ("Infra",      8000.0,  9500.0),
        ].iter().enumerate() {
            let row = (i + 1) as u32;
            ws2.write_string(row, 0, *cat).unwrap();
            ws2.write_number(row, 1, *q1).unwrap();
            ws2.write_number(row, 2, *q2).unwrap();
        }

        let ws3 = wb.add_worksheet();
        ws3.set_name("Summary").unwrap();
        ws3.write_string(0, 0, "Metric").unwrap();
        ws3.write_string(0, 1, "Value").unwrap();
        ws3.write_string(1, 0, "Total Revenue").unwrap();
        ws3.write_number(1, 1, 616000.0).unwrap();
        ws3.write_string(2, 0, "Total Expenses").unwrap();
        ws3.write_number(2, 1, 216500.0).unwrap();
        ws3.write_string(3, 0, "Net").unwrap();
        ws3.write_number(3, 1, 399500.0).unwrap();

        wb.save(dir.join("budget.xlsx")).unwrap();
        println!("Created demo/budget.xlsx");
    }

    // --- old.xlsx & new.xlsx: diff demo pair ---
    {
        let mut wb = Workbook::new();
        let ws = wb.add_worksheet();
        ws.set_name("Sheet1").unwrap();
        ws.write_string(0, 0, "ID").unwrap();
        ws.write_string(0, 1, "Name").unwrap();
        ws.write_string(0, 2, "Department").unwrap();
        ws.write_string(0, 3, "Salary").unwrap();
        for (i, (id, name, dept, salary)) in [
            (1, "Alice",   "Engineering", 95000.0),
            (2, "Bob",     "Marketing",   72000.0),
            (3, "Charlie", "Engineering", 88000.0),
            (4, "Dana",    "Sales",       65000.0),
        ].iter().enumerate() {
            let row = (i + 1) as u32;
            ws.write_number(row, 0, *id as f64).unwrap();
            ws.write_string(row, 1, *name).unwrap();
            ws.write_string(row, 2, *dept).unwrap();
            ws.write_number(row, 3, *salary).unwrap();
        }
        wb.save(dir.join("old.xlsx")).unwrap();
        println!("Created demo/old.xlsx");

        let mut wb = Workbook::new();
        let ws = wb.add_worksheet();
        ws.set_name("Sheet1").unwrap();
        ws.write_string(0, 0, "ID").unwrap();
        ws.write_string(0, 1, "Name").unwrap();
        ws.write_string(0, 2, "Department").unwrap();
        ws.write_string(0, 3, "Salary").unwrap();
        for (i, (id, name, dept, salary)) in [
            (1, "Alice", "Engineering", 98000.0),
            (2, "Bob",   "Design",      75000.0),
            (4, "Dana",  "Sales",       65000.0),
            (5, "Eve",   "Marketing",   70000.0),
        ].iter().enumerate() {
            let row = (i + 1) as u32;
            ws.write_number(row, 0, *id as f64).unwrap();
            ws.write_string(row, 1, *name).unwrap();
            ws.write_string(row, 2, *dept).unwrap();
            ws.write_number(row, 3, *salary).unwrap();
        }
        wb.save(dir.join("new.xlsx")).unwrap();
        println!("Created demo/new.xlsx");
    }
}
