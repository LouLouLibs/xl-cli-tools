use anyhow::{Context, Result};
use calamine::{open_workbook_auto, Data, Reader};
use polars::prelude::*;
use std::path::Path;

#[derive(Debug, Clone, Copy, PartialEq)]
enum InferredType {
    Int,
    Float,
    String,
    Bool,
    DateTime,
    Empty,
}

pub fn read_sheet(path: &Path, sheet_name: &str) -> Result<DataFrame> {
    let mut workbook = open_workbook_auto(path)
        .with_context(|| format!("Cannot open workbook: {}", path.display()))?;
    let range = workbook
        .worksheet_range(sheet_name)
        .with_context(|| format!("Cannot read sheet: {sheet_name}"))?;
    range_to_dataframe(&range)
}

pub fn range_to_dataframe(range: &calamine::Range<Data>) -> Result<DataFrame> {
    let (total_rows, cols) = range.get_size();
    if total_rows == 0 || cols == 0 {
        return Ok(DataFrame::default());
    }

    let rows: Vec<&[Data]> = range.rows().collect();

    // First row = headers
    let headers: Vec<String> = rows[0]
        .iter()
        .enumerate()
        .map(|(i, cell)| match cell {
            Data::String(s) => s.clone(),
            _ => format!("column_{i}"),
        })
        .collect();

    if total_rows == 1 {
        // Header only, no data
        let series: Vec<Column> = headers
            .iter()
            .map(|name| {
                Series::new_empty(PlSmallStr::from(name.as_str()), &DataType::Null).into_column()
            })
            .collect();
        return DataFrame::new(series).map_err(Into::into);
    }

    let data_rows = &rows[1..];
    let mut columns: Vec<Column> = Vec::with_capacity(cols);

    for col_idx in 0..cols {
        let cells: Vec<&Data> = data_rows
            .iter()
            .map(|row| {
                if col_idx < row.len() {
                    &row[col_idx]
                } else {
                    &Data::Empty
                }
            })
            .collect();

        let col_type = infer_column_type(&cells);
        let series = build_series(&headers[col_idx], &cells, col_type)?;
        columns.push(series.into_column());
    }

    DataFrame::new(columns).map_err(Into::into)
}

fn infer_column_type(cells: &[&Data]) -> InferredType {
    let mut has_int = false;
    let mut has_float = false;
    let mut has_string = false;
    let mut has_bool = false;
    let mut has_datetime = false;
    let mut all_empty = true;

    for cell in cells {
        match cell {
            Data::Empty => {}
            Data::String(_) | Data::DateTimeIso(_) | Data::DurationIso(_) => {
                has_string = true;
                all_empty = false;
            }
            Data::Float(_) => {
                has_float = true;
                all_empty = false;
            }
            Data::Int(_) => {
                has_int = true;
                all_empty = false;
            }
            Data::Bool(_) => {
                has_bool = true;
                all_empty = false;
            }
            Data::DateTime(_) => {
                has_datetime = true;
                all_empty = false;
            }
            Data::Error(_) => {
                has_string = true;
                all_empty = false;
            }
        }
    }

    if all_empty {
        return InferredType::Empty;
    }
    // String trumps everything
    if has_string {
        return InferredType::String;
    }
    // DateTime only if all non-empty cells are datetime
    if has_datetime && !has_int && !has_float && !has_bool {
        return InferredType::DateTime;
    }
    // Bool only if all non-empty cells are bool
    if has_bool && !has_int && !has_float && !has_datetime {
        return InferredType::Bool;
    }
    // Float if any float or mix of int/float
    if has_float {
        return InferredType::Float;
    }
    if has_int {
        return InferredType::Int;
    }
    // Fallback: mixed datetime/bool/etc → string
    InferredType::String
}

fn build_series(name: &str, cells: &[&Data], col_type: InferredType) -> Result<Series> {
    let plname = PlSmallStr::from(name);
    match col_type {
        InferredType::Int => {
            let values: Vec<Option<i64>> = cells
                .iter()
                .map(|cell| match cell {
                    Data::Int(v) => Some(*v),
                    Data::Empty => None,
                    _ => None,
                })
                .collect();
            Ok(Series::new(plname, &values))
        }
        InferredType::Float => {
            let values: Vec<Option<f64>> = cells
                .iter()
                .map(|cell| match cell {
                    Data::Float(v) => Some(*v),
                    Data::Int(v) => Some(*v as f64),
                    Data::Empty => None,
                    _ => None,
                })
                .collect();
            Ok(Series::new(plname, &values))
        }
        InferredType::Bool => {
            let values: Vec<Option<bool>> = cells
                .iter()
                .map(|cell| match cell {
                    Data::Bool(v) => Some(*v),
                    Data::Empty => None,
                    _ => None,
                })
                .collect();
            Ok(Series::new(plname, &values))
        }
        InferredType::DateTime => {
            // calamine ExcelDateTime wraps a serial date float (days since 1899-12-30)
            // Convert to milliseconds since Unix epoch for polars
            let values: Vec<Option<i64>> = cells
                .iter()
                .map(|cell| match cell {
                    Data::DateTime(v) => {
                        let serial = v.as_f64();
                        // Excel epoch: 1899-12-30 = -25569 days from Unix epoch
                        let days_from_unix = serial - 25569.0;
                        let ms = (days_from_unix * 86_400_000.0) as i64;
                        Some(ms)
                    }
                    Data::Empty => None,
                    _ => None,
                })
                .collect();
            let series = Series::new(plname, &values);
            Ok(series.cast(&DataType::Datetime(TimeUnit::Milliseconds, None))?)
        }
        InferredType::String | InferredType::Empty => {
            let values: Vec<Option<String>> = cells
                .iter()
                .map(|cell| match cell {
                    Data::String(s) => Some(s.clone()),
                    Data::Float(v) => Some(v.to_string()),
                    Data::Int(v) => Some(v.to_string()),
                    Data::Bool(v) => Some(v.to_string()),
                    Data::DateTime(v) => Some(v.as_f64().to_string()),
                    Data::Error(e) => Some(format!("{e:?}")),
                    Data::DateTimeIso(s) | Data::DurationIso(s) => Some(s.clone()),
                    Data::Empty => None,
                })
                .collect();
            Ok(Series::new(plname, &values))
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use rust_xlsxwriter::Workbook;
    use tempfile::NamedTempFile;

    /// Create a simple xlsx with mixed types: string, float, float (int-like), bool
    fn create_simple(path: &std::path::Path) {
        let mut wb = Workbook::new();
        let ws = wb.add_worksheet().set_name("Data").unwrap();
        ws.write_string(0, 0, "name").unwrap();
        ws.write_string(0, 1, "amount").unwrap();
        ws.write_string(0, 2, "count").unwrap();
        ws.write_string(0, 3, "active").unwrap();

        let names = ["Alice", "Bob", "Charlie", "Diana", "Eve"];
        let amounts = [100.50, 200.75, 300.00, 400.25, 500.00];
        let counts = [10.0, 20.0, 30.0, 40.0, 50.0];
        let active = [true, false, true, false, true];

        for i in 0..5u32 {
            ws.write_string(i + 1, 0, names[i as usize]).unwrap();
            ws.write_number(i + 1, 1, amounts[i as usize]).unwrap();
            ws.write_number(i + 1, 2, counts[i as usize]).unwrap();
            ws.write_boolean(i + 1, 3, active[i as usize]).unwrap();
        }
        wb.save(path).unwrap();
    }

    /// Create xlsx with header row only, no data
    fn create_empty_data(path: &std::path::Path) {
        let mut wb = Workbook::new();
        let ws = wb.add_worksheet().set_name("Empty").unwrap();
        ws.write_string(0, 0, "col_a").unwrap();
        ws.write_string(0, 1, "col_b").unwrap();
        wb.save(path).unwrap();
    }

    /// Create xlsx with a completely empty sheet
    fn create_empty_sheet(path: &std::path::Path) {
        let mut wb = Workbook::new();
        wb.add_worksheet().set_name("Blank").unwrap();
        wb.save(path).unwrap();
    }

    #[test]
    fn test_infer_int_column() {
        let cells = vec![
            &Data::Int(1),
            &Data::Int(2),
            &Data::Empty,
            &Data::Int(4),
        ];
        assert_eq!(infer_column_type(&cells), InferredType::Int);
    }

    #[test]
    fn test_infer_float_when_mixed_int_float() {
        let cells = vec![
            &Data::Int(1),
            &Data::Float(2.5),
            &Data::Int(3),
        ];
        assert_eq!(infer_column_type(&cells), InferredType::Float);
    }

    #[test]
    fn test_infer_string_trumps_all() {
        let s = Data::String("hello".to_string());
        let cells: Vec<&Data> = vec![
            &Data::Int(1),
            &s,
            &Data::Float(3.0),
        ];
        assert_eq!(infer_column_type(&cells), InferredType::String);
    }

    #[test]
    fn test_infer_empty_column() {
        let cells: Vec<&Data> = vec![&Data::Empty, &Data::Empty];
        assert_eq!(infer_column_type(&cells), InferredType::Empty);
    }

    #[test]
    fn test_infer_bool_column() {
        let cells = vec![
            &Data::Bool(true),
            &Data::Bool(false),
            &Data::Empty,
        ];
        assert_eq!(infer_column_type(&cells), InferredType::Bool);
    }

    #[test]
    fn test_empty_range() {
        let range: calamine::Range<Data> = Default::default();
        let df = range_to_dataframe(&range).unwrap();
        assert_eq!(df.height(), 0);
        assert_eq!(df.width(), 0);
    }

    #[test]
    fn test_range_to_dataframe_basic() {
        let tmp = NamedTempFile::with_suffix(".xlsx").unwrap();
        create_simple(tmp.path());

        let df = read_sheet(tmp.path(), "Data").unwrap();
        assert_eq!(df.height(), 5);
        assert_eq!(df.width(), 4);

        let names: Vec<String> = df
            .get_column_names()
            .iter()
            .map(|s| s.to_string())
            .collect();
        assert_eq!(names, vec!["name", "amount", "count", "active"]);
    }

    #[test]
    fn test_read_sheet_types() {
        let tmp = NamedTempFile::with_suffix(".xlsx").unwrap();
        create_simple(tmp.path());

        let df = read_sheet(tmp.path(), "Data").unwrap();

        // name column should be String
        assert_eq!(df.column("name").unwrap().dtype(), &DataType::String);

        // amount column should be Float64
        assert_eq!(df.column("amount").unwrap().dtype(), &DataType::Float64);

        // active column should be Boolean
        assert_eq!(df.column("active").unwrap().dtype(), &DataType::Boolean);
    }

    #[test]
    fn test_read_sheet_header_only() {
        let tmp = NamedTempFile::with_suffix(".xlsx").unwrap();
        create_empty_data(tmp.path());

        let df = read_sheet(tmp.path(), "Empty").unwrap();
        assert_eq!(df.height(), 0);
        assert_eq!(df.width(), 2);
        let names: Vec<String> = df
            .get_column_names()
            .iter()
            .map(|s| s.to_string())
            .collect();
        assert_eq!(names, vec!["col_a", "col_b"]);
    }

    #[test]
    fn test_read_sheet_empty_sheet() {
        let tmp = NamedTempFile::with_suffix(".xlsx").unwrap();
        create_empty_sheet(tmp.path());

        let df = read_sheet(tmp.path(), "Blank").unwrap();
        assert_eq!(df.height(), 0);
        assert_eq!(df.width(), 0);
    }
}
