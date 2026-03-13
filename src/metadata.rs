use anyhow::{Context, Result};
use calamine::{open_workbook_auto, Reader};
use std::path::Path;

/// Info about a single sheet (without loading data).
#[derive(Debug, Clone)]
pub struct SheetInfo {
    pub name: String,
    pub rows: usize, // total rows including header
    pub cols: usize,
}

/// Info about the whole workbook file.
#[derive(Debug)]
pub struct FileInfo {
    pub file_size: u64,
    pub sheets: Vec<SheetInfo>,
}

/// Read metadata: file size, sheet names, and dimensions.
pub fn read_file_info(path: &Path) -> Result<FileInfo> {
    let file_size = std::fs::metadata(path)
        .with_context(|| format!("Cannot read file: {}", path.display()))?
        .len();

    let ext = path
        .extension()
        .and_then(|e| e.to_str())
        .map(|e| e.to_lowercase());

    match ext.as_deref() {
        Some("xlsx") | Some("xls") | Some("xlsb") | Some("xlsm") => {}
        Some(other) => anyhow::bail!("Expected .xls or .xlsx file, got: .{other}"),
        None => anyhow::bail!("Expected .xls or .xlsx file, got: no extension"),
    }

    let mut workbook = open_workbook_auto(path)
        .with_context(|| format!("Cannot open workbook: {}", path.display()))?;

    let sheet_names: Vec<String> = workbook.sheet_names().to_vec();
    let mut sheets = Vec::new();

    for name in &sheet_names {
        let range = workbook
            .worksheet_range(name)
            .with_context(|| format!("Cannot read sheet: {name}"))?;
        let (rows, cols) = range.get_size();
        sheets.push(SheetInfo {
            name: name.clone(),
            rows,
            cols,
        });
    }

    Ok(FileInfo { file_size, sheets })
}

/// Format file size for display: "245 KB", "1.2 MB", etc.
pub fn format_file_size(bytes: u64) -> String {
    if bytes < 1_024 {
        format!("{bytes} B")
    } else if bytes < 1_048_576 {
        format!("{:.0} KB", bytes as f64 / 1_024.0)
    } else if bytes < 1_073_741_824 {
        format!("{:.1} MB", bytes as f64 / 1_048_576.0)
    } else {
        format!("{:.1} GB", bytes as f64 / 1_073_741_824.0)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_format_file_size() {
        assert_eq!(format_file_size(500), "500 B");
        assert_eq!(format_file_size(2_048), "2 KB");
        assert_eq!(format_file_size(1_500_000), "1.4 MB");
    }
}
