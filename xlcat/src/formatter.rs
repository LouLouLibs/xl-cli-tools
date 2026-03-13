use crate::metadata::{format_file_size, FileInfo, SheetInfo};
use polars::prelude::*;

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/// Render the top-level file header.
///
/// ```text
/// # File: report.xlsx (245 KB)
/// # Sheets: 3
/// ```
pub fn format_header(file_name: &str, info: &FileInfo) -> String {
    let size_str = format_file_size(info.file_size);
    let sheet_count = info.sheets.len();
    format!("# File: {file_name} ({size_str})\n# Sheets: {sheet_count}\n")
}

/// Render the schema block for a single sheet.
///
/// ```text
/// ## Sheet: Revenue (1240 rows x 8 cols)
///
/// | Column | Type |
/// |--------|------|
/// | date   | Date |
/// ...
/// ```
pub fn format_schema(sheet: &SheetInfo, df: &DataFrame) -> String {
    let data_rows = if sheet.rows == 0 { 0 } else { sheet.rows - 1 };
    let cols = sheet.cols;

    let mut out = format!(
        "## Sheet: {} ({} rows x {} cols)\n\n",
        sheet.name, data_rows, cols
    );
    out.push_str("| Column | Type |\n");
    out.push_str("|--------|------|\n");

    for col in df.get_columns() {
        let name = col.name();
        let dtype = format_dtype(col.dtype());
        out.push_str(&format!("| {name} | {dtype} |\n"));
    }

    out
}

/// Render the multi-sheet listing (header + schema for each + prompt).
pub fn format_sheet_listing(
    file_name: &str,
    info: &FileInfo,
    schemas: &[(&SheetInfo, DataFrame)],
) -> String {
    let mut out = format_header(file_name, info);
    out.push('\n');

    for (sheet, df) in schemas {
        if sheet.rows == 0 && sheet.cols == 0 {
            out.push_str(&format_empty_sheet(sheet));
        } else {
            out.push_str(&format_schema(sheet, df));
        }
        out.push('\n');
    }

    out.push_str("Use --sheet <name> to view a specific sheet.\n");
    out
}

/// Render the full DataFrame as a markdown table.
pub fn format_data_table(df: &DataFrame) -> String {
    let columns = df.get_columns();
    let n_rows = df.height();

    // Header row
    let mut out = String::new();
    out.push('|');
    for col in columns {
        out.push_str(&format!(" {} |", col.name()));
    }
    out.push('\n');

    // Separator
    out.push('|');
    for _ in columns {
        out.push_str("---|");
    }
    out.push('\n');

    // Data rows
    for row_idx in 0..n_rows {
        out.push('|');
        for col in columns {
            let cell = format_cell(col, row_idx);
            out.push_str(&format!(" {cell} |"));
        }
        out.push('\n');
    }

    out
}

/// Render head / tail view of a DataFrame.
///
/// If total rows <= head_n + tail_n, shows all rows.
/// Otherwise shows first head_n rows, an omission line, then last tail_n rows.
pub fn format_head_tail(df: &DataFrame, head_n: usize, tail_n: usize) -> String {
    let total = df.height();
    if total <= head_n + tail_n {
        return format_data_table(df);
    }

    let head_df = df.head(Some(head_n));
    let tail_df = df.tail(Some(tail_n));
    let omitted = total - head_n - tail_n;

    let mut out = format_data_table(&head_df);
    out.push_str(&format!("... ({omitted} rows omitted) ...\n"));
    // Append tail rows without repeating the header
    let tail_table = format_data_table(&tail_df);
    // Skip header line + separator line of the tail table
    let tail_body = skip_table_header(&tail_table);
    out.push_str(tail_body);

    out
}

/// Render DataFrame as CSV.
pub fn format_csv(df: &DataFrame) -> String {
    let mut buf: Vec<u8> = Vec::new();
    // CsvWriter is available via the "csv" feature (polars 0.46)
    if CsvWriter::new(&mut buf)
        .finish(&mut df.clone())
        .is_ok()
    {
        return String::from_utf8(buf).unwrap_or_else(|_| csv_fallback(df));
    }
    csv_fallback(df)
}

/// Render a message for an empty or header-only sheet.
pub fn format_empty_sheet(sheet: &SheetInfo) -> String {
    if sheet.rows == 0 && sheet.cols == 0 {
        format!("## Sheet: {} (empty)\n", sheet.name)
    } else {
        format!("## Sheet: {} (no data rows)\n", sheet.name)
    }
}

/// Render summary statistics for each column as a markdown table.
///
/// Stats are rows, columns are DataFrame columns:
/// | stat | col1 | col2 | ... |
/// |------|------|------|-----|
/// | count | ... | ... | ... |
/// ...
pub fn format_describe(df: &DataFrame) -> String {
    let columns = df.get_columns();
    let stats = ["count", "null_count", "mean", "std", "min", "max", "median", "unique"];

    // Header row
    let mut out = String::from("| stat |");
    for col in columns {
        out.push_str(&format!(" {} |", col.name()));
    }
    out.push('\n');

    // Separator
    out.push_str("|------|");
    for _ in columns {
        out.push_str("---|");
    }
    out.push('\n');

    // Stat rows
    for stat in &stats {
        out.push_str(&format!("| {stat} |"));
        for col in columns {
            let val = compute_stat(col, stat);
            out.push_str(&format!(" {val} |"));
        }
        out.push('\n');
    }

    out
}

fn compute_stat(col: &Column, stat: &str) -> String {
    let series = col.as_materialized_series();
    match stat {
        "count" => series.len().to_string(),
        "null_count" => series.null_count().to_string(),
        "mean" => {
            if is_numeric(series.dtype()) {
                series.mean().map(|v| format!("{v:.4}")).unwrap_or_else(|| "-".into())
            } else {
                "-".into()
            }
        }
        "std" => {
            if is_numeric(series.dtype()) {
                series.std(1).map(|v| format!("{v:.4}")).unwrap_or_else(|| "-".into())
            } else {
                "-".into()
            }
        }
        "min" => {
            if is_numeric(series.dtype()) {
                match series.min_reduce() {
                    Ok(v) => v.value().to_string(),
                    Err(_) => "-".into(),
                }
            } else {
                "-".into()
            }
        }
        "max" => {
            if is_numeric(series.dtype()) {
                match series.max_reduce() {
                    Ok(v) => v.value().to_string(),
                    Err(_) => "-".into(),
                }
            } else {
                "-".into()
            }
        }
        "median" => {
            if is_numeric(series.dtype()) {
                series.median().map(|v| format!("{v:.4}")).unwrap_or_else(|| "-".into())
            } else {
                "-".into()
            }
        }
        "unique" => match series.n_unique() {
            Ok(n) => n.to_string(),
            Err(_) => "-".into(),
        },
        _ => "-".into(),
    }
}

fn is_numeric(dtype: &DataType) -> bool {
    matches!(
        dtype,
        DataType::Int8
            | DataType::Int16
            | DataType::Int32
            | DataType::Int64
            | DataType::UInt8
            | DataType::UInt16
            | DataType::UInt32
            | DataType::UInt64
            | DataType::Float32
            | DataType::Float64
    )
}

// ---------------------------------------------------------------------------
// Private helpers
// ---------------------------------------------------------------------------

/// Format a single cell value for markdown display.
fn format_cell(col: &Column, idx: usize) -> String {
    match col.get(idx) {
        Ok(AnyValue::Null) | Err(_) => String::new(),
        Ok(v) => format_any_value(&v),
    }
}

/// Convert an AnyValue to its display string.
fn format_any_value(v: &AnyValue) -> String {
    match v {
        AnyValue::Null => String::new(),
        AnyValue::Boolean(b) => b.to_string(),
        AnyValue::Int8(n) => n.to_string(),
        AnyValue::Int16(n) => n.to_string(),
        AnyValue::Int32(n) => n.to_string(),
        AnyValue::Int64(n) => n.to_string(),
        AnyValue::UInt8(n) => n.to_string(),
        AnyValue::UInt16(n) => n.to_string(),
        AnyValue::UInt32(n) => n.to_string(),
        AnyValue::UInt64(n) => n.to_string(),
        AnyValue::Float32(f) => f.to_string(),
        AnyValue::Float64(f) => f.to_string(),
        AnyValue::String(s) => s.to_string(),
        AnyValue::StringOwned(s) => s.to_string(),
        other => format!("{other}"),
    }
}

/// Map a polars DataType to a human-readable label.
fn format_dtype(dtype: &DataType) -> &'static str {
    match dtype {
        DataType::Boolean => "Boolean",
        DataType::Int8 | DataType::Int16 | DataType::Int32 | DataType::Int64 => "Int",
        DataType::UInt8 | DataType::UInt16 | DataType::UInt32 | DataType::UInt64 => "UInt",
        DataType::Float32 | DataType::Float64 => "Float",
        DataType::String => "String",
        DataType::Date => "Date",
        DataType::Datetime(_, _) => "Datetime",
        DataType::Duration(_) => "Duration",
        DataType::Time => "Time",
        DataType::Null => "Null",
        _ => "Other",
    }
}

/// Skip the first two lines (header + separator) of a markdown table string.
fn skip_table_header(table: &str) -> &str {
    let mut newlines = 0;
    let mut pos = 0;
    for (i, ch) in table.char_indices() {
        if ch == '\n' {
            newlines += 1;
            pos = i + 1;
            if newlines == 2 {
                return &table[pos..];
            }
        }
    }
    // Fewer than 2 newlines — return empty
    ""
}

/// Manual CSV fallback if CsvWriter is unavailable.
fn csv_fallback(df: &DataFrame) -> String {
    let columns = df.get_columns();
    let n_rows = df.height();

    let mut out = String::new();

    // Header
    let header: Vec<String> = columns.iter().map(|c| c.name().to_string()).collect();
    out.push_str(&header.join(","));
    out.push('\n');

    // Rows
    for row_idx in 0..n_rows {
        let row: Vec<String> = columns
            .iter()
            .map(|col| {
                let cell = format_cell(col, row_idx);
                // Quote cells containing commas or quotes
                if cell.contains(',') || cell.contains('"') || cell.contains('\n') {
                    format!("\"{}\"", cell.replace('"', "\"\""))
                } else {
                    cell
                }
            })
            .collect();
        out.push_str(&row.join(","));
        out.push('\n');
    }

    out
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

#[cfg(test)]
mod tests {
    use super::*;
    use crate::metadata::{FileInfo, SheetInfo};

    #[test]
    fn test_format_header() {
        let info = FileInfo {
            file_size: 250_000,
            sheets: vec![SheetInfo {
                name: "Sheet1".into(),
                rows: 100,
                cols: 5,
            }],
        };
        let out = format_header("test.xlsx", &info);
        assert!(out.contains("# File: test.xlsx (244 KB)"));
        assert!(out.contains("# Sheets: 1"));
    }

    #[test]
    fn test_format_data_table() {
        let s1 = Series::new("name".into(), &["Alice", "Bob"]);
        let s2 = Series::new("value".into(), &[100i64, 200]);
        let df = DataFrame::new(vec![s1.into_column(), s2.into_column()]).unwrap();
        let out = format_data_table(&df);
        assert!(out.contains("| name | value |"));
        assert!(out.contains("| Alice | 100 |"));
    }

    #[test]
    fn test_format_head_tail_small() {
        let s = Series::new("x".into(), &[1i64, 2, 3]);
        let df = DataFrame::new(vec![s.into_column()]).unwrap();
        let out = format_head_tail(&df, 25, 25);
        assert!(!out.contains("omitted"));
        assert!(out.contains("| 1 |"));
        assert!(out.contains("| 3 |"));
    }

    #[test]
    fn test_format_head_tail_large() {
        // 60 rows, head=25 tail=25 → 10 omitted
        let values: Vec<i64> = (1..=60).collect();
        let s = Series::new("n".into(), values.as_slice());
        let df = DataFrame::new(vec![s.into_column()]).unwrap();
        let out = format_head_tail(&df, 25, 25);
        assert!(out.contains("(10 rows omitted)"));
        assert!(out.contains("| 1 |"));
        assert!(out.contains("| 25 |"));
        assert!(out.contains("| 36 |"));
        assert!(out.contains("| 60 |"));
    }

    #[test]
    fn test_format_schema() {
        let sheet = SheetInfo {
            name: "Revenue".into(),
            rows: 11, // 1 header + 10 data
            cols: 2,
        };
        let s1 = Series::new("date".into(), &["2024-01-01", "2024-01-02"]);
        let s2 = Series::new("amount".into(), &[1.0f64, 2.0]);
        let df = DataFrame::new(vec![s1.into_column(), s2.into_column()]).unwrap();
        let out = format_schema(&sheet, &df);
        assert!(out.contains("## Sheet: Revenue (10 rows x 2 cols)"));
        assert!(out.contains("| date |"));
        assert!(out.contains("| amount |"));
        assert!(out.contains("String"));
        assert!(out.contains("Float"));
    }

    #[test]
    fn test_format_empty_sheet_completely_empty() {
        let sheet = SheetInfo { name: "Blank".into(), rows: 0, cols: 0 };
        let out = format_empty_sheet(&sheet);
        assert!(out.contains("(empty)"));
    }

    #[test]
    fn test_format_empty_sheet_header_only() {
        let sheet = SheetInfo { name: "Headers".into(), rows: 1, cols: 3 };
        let out = format_empty_sheet(&sheet);
        assert!(out.contains("(no data rows)"));
    }

    #[test]
    fn test_format_csv() {
        let s1 = Series::new("a".into(), &["hello", "world"]);
        let s2 = Series::new("b".into(), &[1i64, 2]);
        let df = DataFrame::new(vec![s1.into_column(), s2.into_column()]).unwrap();
        let out = format_csv(&df);
        assert!(out.contains("a,b"));
        assert!(out.contains("hello"));
        assert!(out.contains("world"));
    }

    #[test]
    fn test_format_describe() {
        let s_name = Series::new("name".into(), &["Alice", "Bob", "Carol"]);
        let s_val = Series::new("value".into(), &[10i64, 20, 30]);
        let df = DataFrame::new(vec![s_name.into_column(), s_val.into_column()]).unwrap();
        let out = format_describe(&df);
        // Header row contains stat and column names
        assert!(out.contains("| stat |"));
        assert!(out.contains("| name |"));
        assert!(out.contains("| value |"));
        // All stat rows are present
        assert!(out.contains("| count |"));
        assert!(out.contains("| null_count |"));
        assert!(out.contains("| mean |"));
        assert!(out.contains("| std |"));
        assert!(out.contains("| min |"));
        assert!(out.contains("| max |"));
        assert!(out.contains("| median |"));
        assert!(out.contains("| unique |"));
        // Non-numeric column shows "-" for mean
        assert!(out.contains("| mean | - |"));
        // Numeric column has a numeric mean value (not "-")
        // count=3 for both
        assert!(out.contains("| count | 3 | 3 |"));
        // unique=3 for both
        assert!(out.contains("| unique | 3 | 3 |"));
    }
}
