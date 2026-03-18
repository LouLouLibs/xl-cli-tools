// Diff engine for comparing two Excel sheets.

use polars::prelude::*;

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
