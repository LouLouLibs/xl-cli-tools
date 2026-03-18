// Diff engine for comparing two Excel sheets.

use anyhow::{Result, bail};
use polars::prelude::*;
use std::collections::HashMap;

use crate::formatter;

/// Source file and sheet metadata for display.
#[derive(Debug, Clone)]
pub struct SheetSource {
    pub file_name: String,
    pub sheet_name: String,
}

/// A single row from an added or removed set.
#[derive(Debug, Clone)]
pub struct DiffRow {
    pub values: Vec<String>,
}

/// A change in a single cell.
#[derive(Debug, Clone)]
pub struct CellChange {
    pub column: String,
    pub old_value: String,
    pub new_value: String,
}

/// A row present in both files with cell-level differences.
#[derive(Debug, Clone)]
pub struct ModifiedRow {
    pub key: Vec<String>,
    pub changes: Vec<CellChange>,
}

/// Result of comparing two sheets.
#[derive(Debug, Clone)]
pub struct DiffResult {
    pub headers: Vec<String>,
    pub key_columns: Vec<String>,
    pub added: Vec<DiffRow>,
    pub removed: Vec<DiffRow>,
    pub modified: Vec<ModifiedRow>,
    pub source_a: SheetSource,
    pub source_b: SheetSource,
}

impl DiffResult {
    pub fn has_differences(&self) -> bool {
        !self.added.is_empty() || !self.removed.is_empty() || !self.modified.is_empty()
    }
}

/// Options controlling how the diff is performed.
#[derive(Debug, Clone, Default)]
pub struct DiffOptions {
    pub key_columns: Vec<String>,
    pub tolerance: Option<f64>,
}

// ---------------------------------------------------------------------------
// Helper functions
// ---------------------------------------------------------------------------

/// Format a cell value for display. Returns empty string for null.
fn cell_to_string(col: &Column, idx: usize) -> String {
    match col.get(idx) {
        Ok(AnyValue::Null) | Err(_) => String::new(),
        Ok(v) => formatter::format_any_value(&v),
    }
}

/// Format a cell value for hashing. Uses a sentinel for null so that null
/// and empty string produce different keys.
fn cell_to_key_part(col: &Column, idx: usize) -> String {
    match col.get(idx) {
        Ok(AnyValue::Null) | Err(_) => "\x01NULL\x01".to_string(),
        Ok(v) => formatter::format_any_value(&v),
    }
}

/// Build a string key for an entire row by joining all column values.
fn row_to_key(df: &DataFrame, row_idx: usize) -> String {
    df.get_columns()
        .iter()
        .map(|col| cell_to_key_part(col, row_idx))
        .collect::<Vec<_>>()
        .join("\0")
}

/// Collect display values for every column in a row.
fn row_to_strings(df: &DataFrame, row_idx: usize) -> Vec<String> {
    df.get_columns()
        .iter()
        .map(|col| cell_to_string(col, row_idx))
        .collect()
}

// ---------------------------------------------------------------------------
// Positional diff
// ---------------------------------------------------------------------------

/// Compare two DataFrames positionally (no key columns).
///
/// Uses multiset comparison: each unique row is tracked by frequency.
/// Rows present in A but not (or fewer times) in B are "removed";
/// rows present in B but not (or fewer times) in A are "added".
pub fn diff_positional(
    df_a: &DataFrame,
    df_b: &DataFrame,
    _opts: &DiffOptions,
    source_a: SheetSource,
    source_b: SheetSource,
) -> Result<DiffResult> {
    // Determine headers — use the longer header set.
    let headers_a: Vec<String> = df_a.get_column_names().iter().map(|s| s.to_string()).collect();
    let headers_b: Vec<String> = df_b.get_column_names().iter().map(|s| s.to_string()).collect();

    let headers = if headers_b.len() > headers_a.len() {
        if headers_a.len() != headers_b.len() {
            eprintln!(
                "Warning: column count differs ({} vs {}), using wider header set",
                headers_a.len(),
                headers_b.len()
            );
        }
        headers_b.clone()
    } else {
        if headers_a.len() != headers_b.len() {
            eprintln!(
                "Warning: column count differs ({} vs {}), using wider header set",
                headers_a.len(),
                headers_b.len()
            );
        }
        headers_a.clone()
    };

    let num_headers = headers.len();

    // Build frequency maps: key → list of row indices (so we can consume them).
    let mut freq_a: HashMap<String, Vec<usize>> = HashMap::new();
    for i in 0..df_a.height() {
        let key = row_to_key(df_a, i);
        freq_a.entry(key).or_default().push(i);
    }

    let mut freq_b: HashMap<String, Vec<usize>> = HashMap::new();
    for i in 0..df_b.height() {
        let key = row_to_key(df_b, i);
        freq_b.entry(key).or_default().push(i);
    }

    let mut removed = Vec::new();
    let mut added = Vec::new();

    // Walk A: for each row, try to consume a matching row from B.
    for i in 0..df_a.height() {
        let key = row_to_key(df_a, i);
        let consumed = freq_b
            .get_mut(&key)
            .and_then(|indices| indices.pop())
            .is_some();
        if !consumed {
            let mut vals = row_to_strings(df_a, i);
            vals.resize(num_headers, String::new());
            removed.push(DiffRow { values: vals });
        }
    }

    // Walk B: for each row, try to consume a matching row from A.
    for i in 0..df_b.height() {
        let key = row_to_key(df_b, i);
        let consumed = freq_a
            .get_mut(&key)
            .and_then(|indices| indices.pop())
            .is_some();
        if !consumed {
            let mut vals = row_to_strings(df_b, i);
            vals.resize(num_headers, String::new());
            added.push(DiffRow { values: vals });
        }
    }

    Ok(DiffResult {
        headers,
        key_columns: vec![],
        added,
        removed,
        modified: vec![],
        source_a,
        source_b,
    })
}

// ---------------------------------------------------------------------------
// Key-based diff
// ---------------------------------------------------------------------------

/// A row indexed by its key columns.
struct KeyedRow {
    values: Vec<String>,
    key_values: Vec<String>,
}

/// Build a map from composite key string to KeyedRow for every row in the DataFrame.
fn build_key_map(
    df: &DataFrame,
    key_indices: &[usize],
    columns: &[Column],
) -> HashMap<String, KeyedRow> {
    let mut map = HashMap::new();
    for i in 0..df.height() {
        let key_values: Vec<String> = key_indices
            .iter()
            .map(|&ki| cell_to_string(&columns[ki], i))
            .collect();
        let composite_key = key_values.join("\0");
        let values: Vec<String> = columns.iter().map(|col| cell_to_string(col, i)).collect();
        map.insert(
            composite_key,
            KeyedRow {
                values,
                key_values,
            },
        );
    }
    map
}

/// Warn on stderr when duplicate keys are found.
fn check_duplicate_keys(
    df: &DataFrame,
    key_indices: &[usize],
    columns: &[Column],
    source: &SheetSource,
) {
    let mut seen: HashMap<String, usize> = HashMap::new();
    for i in 0..df.height() {
        let key: String = key_indices
            .iter()
            .map(|&ki| cell_to_string(&columns[ki], i))
            .collect::<Vec<_>>()
            .join("\0");
        let count = seen.entry(key.clone()).or_insert(0);
        *count += 1;
        if *count == 2 {
            let display_key = key.replace('\0', ", ");
            eprintln!(
                "Warning: duplicate key [{}] in {}:{}",
                display_key, source.file_name, source.sheet_name
            );
        }
    }
}

/// Check whether a polars DataType is a float type.
fn is_float_dtype(dt: &DataType) -> bool {
    matches!(dt, DataType::Float32 | DataType::Float64)
}

/// Check whether a polars DataType is an integer type.
fn is_int_dtype(dt: &DataType) -> bool {
    matches!(
        dt,
        DataType::Int8
            | DataType::Int16
            | DataType::Int32
            | DataType::Int64
            | DataType::UInt8
            | DataType::UInt16
            | DataType::UInt32
            | DataType::UInt64
    )
}

/// Compare two string-rendered values with optional numeric tolerance.
///
/// Rules:
/// - NaN == NaN is true.
/// - NaN vs non-NaN is false.
/// - Pure int+int columns use exact comparison (no tolerance applied).
/// - At least one float column applies tolerance.
/// - Otherwise exact string comparison.
fn values_equal_with_tolerance(
    val_a: &str,
    val_b: &str,
    tolerance: f64,
    df_a: &DataFrame,
    df_b: &DataFrame,
    col_name: &str,
) -> bool {
    let parsed_a = val_a.parse::<f64>();
    let parsed_b = val_b.parse::<f64>();

    match (parsed_a, parsed_b) {
        (Ok(a), Ok(b)) => {
            if a.is_nan() && b.is_nan() {
                return true;
            }
            if a.is_nan() || b.is_nan() {
                return false;
            }

            let dt_a = df_a
                .column(col_name)
                .map(|c| c.dtype().clone())
                .unwrap_or(DataType::String);
            let dt_b = df_b
                .column(col_name)
                .map(|c| c.dtype().clone())
                .unwrap_or(DataType::String);

            if is_int_dtype(&dt_a) && is_int_dtype(&dt_b) {
                val_a == val_b
            } else if is_float_dtype(&dt_a) || is_float_dtype(&dt_b) {
                (a - b).abs() <= tolerance
            } else {
                val_a == val_b
            }
        }
        _ => val_a == val_b,
    }
}

/// Compare non-key columns of two keyed rows and return cell-level changes.
fn compare_rows(
    df_a: &DataFrame,
    df_b: &DataFrame,
    headers_a: &[String],
    headers_b: &[String],
    row_a: &KeyedRow,
    row_b: &KeyedRow,
    common_columns: &[String],
    opts: &DiffOptions,
) -> Vec<CellChange> {
    let mut changes = Vec::new();
    for col_name in common_columns {
        let idx_a = headers_a.iter().position(|h| h == col_name);
        let idx_b = headers_b.iter().position(|h| h == col_name);
        let val_a = idx_a
            .map(|i| row_a.values.get(i).cloned().unwrap_or_default())
            .unwrap_or_default();
        let val_b = idx_b
            .map(|i| row_b.values.get(i).cloned().unwrap_or_default())
            .unwrap_or_default();

        let equal = if let Some(tol) = opts.tolerance {
            values_equal_with_tolerance(&val_a, &val_b, tol, df_a, df_b, col_name)
        } else {
            val_a == val_b
        };

        if !equal {
            changes.push(CellChange {
                column: col_name.clone(),
                old_value: val_a,
                new_value: val_b,
            });
        }
    }
    changes
}

/// Compare two DataFrames using key columns.
pub fn diff_keyed(
    df_a: &DataFrame,
    df_b: &DataFrame,
    opts: &DiffOptions,
    source_a: SheetSource,
    source_b: SheetSource,
) -> Result<DiffResult> {
    let columns_a = df_a.get_columns();
    let columns_b = df_b.get_columns();
    let headers_a: Vec<String> = df_a.get_column_names().iter().map(|s| s.to_string()).collect();
    let headers_b: Vec<String> = df_b.get_column_names().iter().map(|s| s.to_string()).collect();

    // Resolve key column indices in both frames.
    let mut key_indices_a = Vec::new();
    let mut key_indices_b = Vec::new();
    for key_col in &opts.key_columns {
        match headers_a.iter().position(|h| h == key_col) {
            Some(idx) => key_indices_a.push(idx),
            None => bail!("Key column '{}' not found in {}", key_col, source_a.file_name),
        }
        match headers_b.iter().position(|h| h == key_col) {
            Some(idx) => key_indices_b.push(idx),
            None => bail!("Key column '{}' not found in {}", key_col, source_b.file_name),
        }
    }

    // Find non-key columns.
    let non_key_a: Vec<String> = headers_a
        .iter()
        .filter(|h| !opts.key_columns.contains(h))
        .cloned()
        .collect();
    let non_key_b: Vec<String> = headers_b
        .iter()
        .filter(|h| !opts.key_columns.contains(h))
        .cloned()
        .collect();

    // Common non-key columns (for modification detection).
    let common_columns: Vec<String> = non_key_a
        .iter()
        .filter(|h| non_key_b.contains(h))
        .cloned()
        .collect();

    // Warn about columns only in one file.
    for col in &non_key_a {
        if !non_key_b.contains(col) {
            eprintln!("Warning: column '{}' only in {}", col, source_a.file_name);
        }
    }
    for col in &non_key_b {
        if !non_key_a.contains(col) {
            eprintln!("Warning: column '{}' only in {}", col, source_b.file_name);
        }
    }

    // Build output headers: key columns + all from A non-key + B-only non-key.
    let mut headers = opts.key_columns.clone();
    headers.extend(non_key_a.iter().cloned());
    for col in &non_key_b {
        if !non_key_a.contains(col) {
            headers.push(col.clone());
        }
    }

    // Check for duplicate keys.
    check_duplicate_keys(df_a, &key_indices_a, columns_a, &source_a);
    check_duplicate_keys(df_b, &key_indices_b, columns_b, &source_b);

    // Build key maps.
    let map_a = build_key_map(df_a, &key_indices_a, columns_a);
    let map_b = build_key_map(df_b, &key_indices_b, columns_b);

    let mut removed = Vec::new();
    let mut added = Vec::new();
    let mut modified = Vec::new();

    // Keys in A but not in B → removed.
    for (composite_key, row_a) in &map_a {
        if !map_b.contains_key(composite_key) {
            let mut vals = Vec::new();
            for h in &headers {
                if let Some(idx) = headers_a.iter().position(|ha| ha == h) {
                    vals.push(row_a.values.get(idx).cloned().unwrap_or_default());
                } else {
                    vals.push(String::new());
                }
            }
            removed.push(DiffRow { values: vals });
        }
    }

    // Keys in B but not in A → added.
    for (composite_key, row_b) in &map_b {
        if !map_a.contains_key(composite_key) {
            let mut vals = Vec::new();
            for h in &headers {
                if let Some(idx) = headers_b.iter().position(|hb| hb == h) {
                    vals.push(row_b.values.get(idx).cloned().unwrap_or_default());
                } else {
                    vals.push(String::new());
                }
            }
            added.push(DiffRow { values: vals });
        }
    }

    // Keys in both → compare for modifications.
    for (composite_key, row_a) in &map_a {
        if let Some(row_b) = map_b.get(composite_key) {
            let changes = compare_rows(
                df_a,
                df_b,
                &headers_a,
                &headers_b,
                row_a,
                row_b,
                &common_columns,
                opts,
            );
            if !changes.is_empty() {
                modified.push(ModifiedRow {
                    key: row_a.key_values.clone(),
                    changes,
                });
            }
        }
    }

    Ok(DiffResult {
        headers,
        key_columns: opts.key_columns.clone(),
        added,
        removed,
        modified,
        source_a,
        source_b,
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    fn test_source_a() -> SheetSource {
        SheetSource {
            file_name: "a.xlsx".into(),
            sheet_name: "Sheet1".into(),
        }
    }

    fn test_source_b() -> SheetSource {
        SheetSource {
            file_name: "b.xlsx".into(),
            sheet_name: "Sheet1".into(),
        }
    }

    // ---- Positional diff tests ----

    #[test]
    fn test_positional_no_diff() {
        let df_a = df! {
            "name" => &["Alice", "Bob"],
            "score" => &[100, 200],
        }
        .unwrap();
        let df_b = df_a.clone();
        let opts = DiffOptions::default();

        let result = diff_positional(&df_a, &df_b, &opts, test_source_a(), test_source_b()).unwrap();

        assert!(!result.has_differences());
        assert!(result.added.is_empty());
        assert!(result.removed.is_empty());
        assert!(result.modified.is_empty());
    }

    #[test]
    fn test_positional_added_removed() {
        let df_a = df! {
            "name" => &["Alice", "Bob"],
        }
        .unwrap();
        let df_b = df! {
            "name" => &["Alice", "Charlie"],
        }
        .unwrap();
        let opts = DiffOptions::default();

        let result = diff_positional(&df_a, &df_b, &opts, test_source_a(), test_source_b()).unwrap();

        assert!(result.has_differences());
        assert_eq!(result.removed.len(), 1);
        assert_eq!(result.removed[0].values, vec!["Bob"]);
        assert_eq!(result.added.len(), 1);
        assert_eq!(result.added[0].values, vec!["Charlie"]);
    }

    #[test]
    fn test_positional_duplicate_rows() {
        let df_a = df! {
            "val" => &["A", "A", "A"],
        }
        .unwrap();
        let df_b = df! {
            "val" => &["A", "A"],
        }
        .unwrap();
        let opts = DiffOptions::default();

        let result = diff_positional(&df_a, &df_b, &opts, test_source_a(), test_source_b()).unwrap();

        assert_eq!(result.removed.len(), 1);
        assert_eq!(result.removed[0].values, vec!["A"]);
        assert!(result.added.is_empty());
    }

    // ---- Key-based diff tests ----

    #[test]
    fn test_keyed_no_diff() {
        let df_a = df! {
            "id" => &[1, 2],
            "name" => &["Alice", "Bob"],
        }
        .unwrap();
        let df_b = df_a.clone();
        let opts = DiffOptions {
            key_columns: vec!["id".into()],
            tolerance: None,
        };

        let result = diff_keyed(&df_a, &df_b, &opts, test_source_a(), test_source_b()).unwrap();

        assert!(!result.has_differences());
    }

    #[test]
    fn test_keyed_added_removed() {
        let df_a = df! {
            "id" => &[1, 2],
            "name" => &["Alice", "Bob"],
        }
        .unwrap();
        let df_b = df! {
            "id" => &[2, 3],
            "name" => &["Bob", "Charlie"],
        }
        .unwrap();
        let opts = DiffOptions {
            key_columns: vec!["id".into()],
            tolerance: None,
        };

        let result = diff_keyed(&df_a, &df_b, &opts, test_source_a(), test_source_b()).unwrap();

        assert_eq!(result.removed.len(), 1);
        assert!(result.removed[0].values.contains(&"1".to_string()));
        assert!(result.removed[0].values.contains(&"Alice".to_string()));

        assert_eq!(result.added.len(), 1);
        assert!(result.added[0].values.contains(&"3".to_string()));
        assert!(result.added[0].values.contains(&"Charlie".to_string()));
    }

    #[test]
    fn test_keyed_modified() {
        let df_a = df! {
            "id" => &[1, 2],
            "score" => &[100, 200],
        }
        .unwrap();
        let df_b = df! {
            "id" => &[1, 2],
            "score" => &[100, 250],
        }
        .unwrap();
        let opts = DiffOptions {
            key_columns: vec!["id".into()],
            tolerance: None,
        };

        let result = diff_keyed(&df_a, &df_b, &opts, test_source_a(), test_source_b()).unwrap();

        assert!(result.added.is_empty());
        assert!(result.removed.is_empty());
        assert_eq!(result.modified.len(), 1);

        let m = &result.modified[0];
        assert_eq!(m.key, vec!["2"]);
        assert_eq!(m.changes.len(), 1);
        assert_eq!(m.changes[0].column, "score");
        assert_eq!(m.changes[0].old_value, "200");
        assert_eq!(m.changes[0].new_value, "250");
    }

    #[test]
    fn test_keyed_composite_key() {
        let df_a = df! {
            "date" => &["2024-01-01", "2024-01-01", "2024-01-02"],
            "ticker" => &["AAPL", "GOOG", "AAPL"],
            "price" => &[150.0, 140.0, 151.0],
        }
        .unwrap();
        let df_b = df! {
            "date" => &["2024-01-01", "2024-01-01", "2024-01-02"],
            "ticker" => &["AAPL", "GOOG", "AAPL"],
            "price" => &[150.0, 142.0, 151.0],
        }
        .unwrap();
        let opts = DiffOptions {
            key_columns: vec!["date".into(), "ticker".into()],
            tolerance: None,
        };

        let result = diff_keyed(&df_a, &df_b, &opts, test_source_a(), test_source_b()).unwrap();

        assert!(result.added.is_empty());
        assert!(result.removed.is_empty());
        assert_eq!(result.modified.len(), 1);

        let m = &result.modified[0];
        assert_eq!(m.key, vec!["2024-01-01", "GOOG"]);
        assert_eq!(m.changes[0].column, "price");
        assert_eq!(m.changes[0].old_value, "140");
        assert_eq!(m.changes[0].new_value, "142");
    }

    // ---- Tolerance tests ----

    #[test]
    fn test_keyed_tolerance_within() {
        let df_a = df! {
            "id" => &[1],
            "price" => &[100.001_f64],
        }
        .unwrap();
        let df_b = df! {
            "id" => &[1],
            "price" => &[100.002_f64],
        }
        .unwrap();
        let opts = DiffOptions {
            key_columns: vec!["id".into()],
            tolerance: Some(0.01),
        };

        let result = diff_keyed(&df_a, &df_b, &opts, test_source_a(), test_source_b()).unwrap();

        assert!(!result.has_differences());
    }

    #[test]
    fn test_keyed_tolerance_exceeded() {
        let df_a = df! {
            "id" => &[1],
            "price" => &[100.0_f64],
        }
        .unwrap();
        let df_b = df! {
            "id" => &[1],
            "price" => &[100.05_f64],
        }
        .unwrap();
        let opts = DiffOptions {
            key_columns: vec!["id".into()],
            tolerance: Some(0.01),
        };

        let result = diff_keyed(&df_a, &df_b, &opts, test_source_a(), test_source_b()).unwrap();

        assert_eq!(result.modified.len(), 1);
        assert_eq!(result.modified[0].changes[0].column, "price");
    }

    #[test]
    fn test_keyed_nan_handling() {
        let df_a = df! {
            "id" => &[1],
            "value" => &[f64::NAN],
        }
        .unwrap();
        let df_b = df! {
            "id" => &[1],
            "value" => &[f64::NAN],
        }
        .unwrap();
        let opts = DiffOptions {
            key_columns: vec!["id".into()],
            tolerance: Some(0.01),
        };

        let result = diff_keyed(&df_a, &df_b, &opts, test_source_a(), test_source_b()).unwrap();

        assert!(!result.has_differences(), "NaN vs NaN should be treated as equal");
    }
}
