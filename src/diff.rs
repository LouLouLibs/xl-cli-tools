// Diff engine for comparing two Excel sheets.

use anyhow::Result;
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
}
