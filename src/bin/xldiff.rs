use std::io::IsTerminal;
use std::path::PathBuf;
use std::process;

use anyhow::{Result, bail};
use clap::Parser;
use serde_json::{Map, Value, json};

use xlcat::diff::{DiffOptions, DiffResult, SheetSource};
use xlcat::filter;
use xlcat::formatter;
use xlcat::metadata;
use xlcat::reader;

#[derive(Parser)]
#[command(
    name = "xldiff",
    about = "Compare two Excel spreadsheets and show differences",
    version
)]
struct Args {
    /// First file (optionally file.xlsx:SheetName)
    file_a: String,

    /// Second file (optionally file.xlsx:SheetName)
    file_b: String,

    /// Key column(s) for matching rows (comma-separated names or letters)
    #[arg(long)]
    key: Option<String>,

    /// Columns to compare (comma-separated names or letters)
    #[arg(long)]
    cols: Option<String>,

    /// Rows to skip before header: single number or "skipA,skipB"
    #[arg(long, default_value = "0")]
    skip: String,

    /// Treat first row as data, not headers (declared for future use)
    #[arg(long)]
    no_header: bool,

    /// Numeric tolerance for float comparisons
    #[arg(long)]
    tolerance: Option<f64>,

    /// Output format: text, markdown, json, csv
    #[arg(long, default_value = "text")]
    format: String,

    /// Disable colored output
    #[arg(long)]
    no_color: bool,
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/// Split a CLI argument like "file.xlsx:Sheet1" into (path, optional sheet).
///
/// Handles Windows drive letters (e.g. C:\file.xlsx) by ignoring a colon
/// at position 1 when followed by `\` or `/`. A trailing colon with nothing
/// after it is treated as no sheet specification.
fn parse_file_arg(arg: &str) -> (PathBuf, Option<String>) {
    // Find the *last* colon — that's the sheet separator.
    if let Some(pos) = arg.rfind(':') {
        // Skip Windows drive letters: colon at position 1 followed by \ or /
        if pos == 1 {
            let after = arg.as_bytes().get(2);
            if after == Some(&b'\\') || after == Some(&b'/') {
                return (PathBuf::from(arg), None);
            }
        }
        let sheet_part = &arg[pos + 1..];
        if sheet_part.is_empty() {
            // Trailing colon — ignore it
            return (PathBuf::from(&arg[..pos]), None);
        }
        (PathBuf::from(&arg[..pos]), Some(sheet_part.to_string()))
    } else {
        (PathBuf::from(arg), None)
    }
}

/// Parse the --skip flag: either "3" (same skip for both) or "3,5".
fn parse_skip(s: &str) -> Result<(usize, usize)> {
    if let Some((a, b)) = s.split_once(',') {
        let skip_a: usize = a
            .trim()
            .parse()
            .map_err(|_| anyhow::anyhow!("invalid skip value: '{}'", a.trim()))?;
        let skip_b: usize = b
            .trim()
            .parse()
            .map_err(|_| anyhow::anyhow!("invalid skip value: '{}'", b.trim()))?;
        Ok((skip_a, skip_b))
    } else {
        let skip: usize = s
            .trim()
            .parse()
            .map_err(|_| anyhow::anyhow!("invalid skip value: '{}'", s.trim()))?;
        Ok((skip, skip))
    }
}

/// Resolve a sheet name: exact match, then 0-based index, then error.
fn resolve_sheet(
    info: &metadata::FileInfo,
    sheet_arg: Option<&str>,
) -> Result<String> {
    match sheet_arg {
        None => info
            .sheets
            .first()
            .map(|s| s.name.clone())
            .ok_or_else(|| anyhow::anyhow!("workbook has no sheets")),
        Some(s) => {
            // Exact name match
            if let Some(sheet) = info.sheets.iter().find(|si| si.name == s) {
                return Ok(sheet.name.clone());
            }
            // 0-based index
            if let Ok(idx) = s.parse::<usize>()
                && let Some(sheet) = info.sheets.get(idx)
            {
                return Ok(sheet.name.clone());
            }
            let names: Vec<_> = info.sheets.iter().map(|si| si.name.as_str()).collect();
            bail!(
                "sheet '{}' not found. Available sheets: {}",
                s,
                names.join(", ")
            )
        }
    }
}

// ---------------------------------------------------------------------------
// Output formatters
// ---------------------------------------------------------------------------

/// Format a row's values inline: `Name: "Alice"  Score: "90"`
fn format_row_inline(headers: &[String], values: &[String]) -> String {
    headers
        .iter()
        .zip(values.iter())
        .map(|(h, v)| format!("{}: \"{}\"", h, v))
        .collect::<Vec<_>>()
        .join("  ")
}

/// Format diff result as colored (or plain) text output.
fn format_text(result: &DiffResult, color: bool) -> String {
    if !result.has_differences() {
        return "No differences found.\n".to_string();
    }

    let (red, green, yellow, reset) = if color {
        ("\x1b[31m", "\x1b[32m", "\x1b[33m", "\x1b[0m")
    } else {
        ("", "", "", "")
    };

    let mut out = String::new();

    // Header
    out.push_str(&format!(
        "--- {} ({})\n+++ {} ({})\n\n",
        result.source_a.sheet_name,
        result.source_a.file_name,
        result.source_b.sheet_name,
        result.source_b.file_name,
    ));

    // Summary
    out.push_str(&format!(
        "Added: {} | Removed: {} | Modified: {}\n\n",
        result.added.len(),
        result.removed.len(),
        result.modified.len(),
    ));

    // Removed rows
    for row in &result.removed {
        out.push_str(&format!(
            "{}- {}{}",
            red,
            format_row_inline(&result.headers, &row.values),
            reset,
        ));
        out.push('\n');
    }

    // Added rows
    for row in &result.added {
        out.push_str(&format!(
            "{}+ {}{}",
            green,
            format_row_inline(&result.headers, &row.values),
            reset,
        ));
        out.push('\n');
    }

    // Modified rows
    for m in &result.modified {
        // Build key display
        let key_display: Vec<String> = result
            .key_columns
            .iter()
            .zip(m.key.iter())
            .map(|(col, val)| format!("{}={}", col, val))
            .collect();
        out.push_str(&format!(
            "{}~ [{}]{}",
            yellow,
            key_display.join(", "),
            reset,
        ));
        out.push('\n');
        for change in &m.changes {
            out.push_str(&format!(
                "    {}: \"{}\" \u{2192} \"{}\"\n",
                change.column, change.old_value, change.new_value,
            ));
        }
    }

    out
}

/// Format diff result as markdown.
fn format_markdown(result: &DiffResult) -> String {
    if !result.has_differences() {
        return "No differences found.\n".to_string();
    }

    let mut out = String::new();

    // Added
    if !result.added.is_empty() {
        out.push_str(&format!("## Added ({})\n\n", result.added.len()));
        let rows: Vec<Vec<String>> = result
            .added
            .iter()
            .map(|r| r.values.clone())
            .collect();
        out.push_str(&formatter::render_table(&result.headers, &rows));
        out.push('\n');
    }

    // Removed
    if !result.removed.is_empty() {
        out.push_str(&format!("## Removed ({})\n\n", result.removed.len()));
        let rows: Vec<Vec<String>> = result
            .removed
            .iter()
            .map(|r| r.values.clone())
            .collect();
        out.push_str(&formatter::render_table(&result.headers, &rows));
        out.push('\n');
    }

    // Modified
    if !result.modified.is_empty() {
        out.push_str(&format!("## Modified ({})\n\n", result.modified.len()));

        let headers = vec![
            "Key".to_string(),
            "Column".to_string(),
            "Old".to_string(),
            "New".to_string(),
        ];
        let mut rows: Vec<Vec<String>> = Vec::new();

        for m in &result.modified {
            let key_display: String = result
                .key_columns
                .iter()
                .zip(m.key.iter())
                .map(|(col, val)| format!("{}={}", col, val))
                .collect::<Vec<_>>()
                .join(", ");

            for (i, change) in m.changes.iter().enumerate() {
                let key_cell = if i == 0 {
                    key_display.clone()
                } else {
                    String::new()
                };
                rows.push(vec![
                    key_cell,
                    change.column.clone(),
                    change.old_value.clone(),
                    change.new_value.clone(),
                ]);
            }
        }

        out.push_str(&formatter::render_table(&headers, &rows));
        out.push('\n');
    }

    out
}

/// Format diff result as JSON.
fn format_json(result: &DiffResult) -> String {
    let added: Vec<Value> = result
        .added
        .iter()
        .map(|row| {
            let mut map = Map::new();
            for (h, v) in result.headers.iter().zip(row.values.iter()) {
                map.insert(h.clone(), Value::String(v.clone()));
            }
            Value::Object(map)
        })
        .collect();

    let removed: Vec<Value> = result
        .removed
        .iter()
        .map(|row| {
            let mut map = Map::new();
            for (h, v) in result.headers.iter().zip(row.values.iter()) {
                map.insert(h.clone(), Value::String(v.clone()));
            }
            Value::Object(map)
        })
        .collect();

    let modified: Vec<Value> = result
        .modified
        .iter()
        .map(|m| {
            let mut key_map = Map::new();
            for (col, val) in result.key_columns.iter().zip(m.key.iter()) {
                key_map.insert(col.clone(), Value::String(val.clone()));
            }
            let changes: Vec<Value> = m
                .changes
                .iter()
                .map(|c| {
                    json!({
                        "column": c.column,
                        "old": c.old_value,
                        "new": c.new_value,
                    })
                })
                .collect();
            json!({
                "key": Value::Object(key_map),
                "changes": changes,
            })
        })
        .collect();

    let output = json!({
        "added": added,
        "removed": removed,
        "modified": modified,
    });

    serde_json::to_string_pretty(&output).unwrap() + "\n"
}

/// Quote a value per RFC 4180: if it contains comma, quote, or newline, wrap
/// in double quotes and escape any internal quotes by doubling them.
fn csv_quote(value: &str) -> String {
    if value.contains(',') || value.contains('"') || value.contains('\n') {
        format!("\"{}\"", value.replace('"', "\"\""))
    } else {
        value.to_string()
    }
}

/// Build a CSV row from a slice of values.
fn csv_row(values: &[String]) -> String {
    values.iter().map(|v| csv_quote(v)).collect::<Vec<_>>().join(",")
}

/// Format diff result as CSV.
///
/// Header: _status, col1, col2, ..., _old_col1, _old_col2, ...
/// Added rows: "added" + values + empty _old_ columns
/// Removed rows: "removed" + values + empty _old_ columns
/// Modified rows: "modified" + new values (key cols + changed new values) + old values in _old_ columns
fn format_csv(result: &DiffResult) -> String {
    let mut out = String::new();

    // Build header
    let mut header_parts: Vec<String> = vec!["_status".to_string()];
    for h in &result.headers {
        header_parts.push(h.clone());
    }
    for h in &result.headers {
        header_parts.push(format!("_old_{}", h));
    }
    out.push_str(&csv_row(&header_parts));
    out.push('\n');

    let empty_cols: Vec<String> = result.headers.iter().map(|_| String::new()).collect();

    // Removed rows
    for row in &result.removed {
        let mut parts: Vec<String> = vec!["removed".to_string()];
        parts.extend(row.values.iter().cloned());
        // Pad if row has fewer values than headers
        while parts.len() < 1 + result.headers.len() {
            parts.push(String::new());
        }
        parts.extend(empty_cols.iter().cloned());
        out.push_str(&csv_row(&parts));
        out.push('\n');
    }

    // Added rows
    for row in &result.added {
        let mut parts: Vec<String> = vec!["added".to_string()];
        parts.extend(row.values.iter().cloned());
        while parts.len() < 1 + result.headers.len() {
            parts.push(String::new());
        }
        parts.extend(empty_cols.iter().cloned());
        out.push_str(&csv_row(&parts));
        out.push('\n');
    }

    // Modified rows
    for m in &result.modified {
        let mut main_cols: Vec<String> = Vec::new();
        let mut old_cols: Vec<String> = Vec::new();

        for h in &result.headers {
            // Check if this is a key column
            if let Some(key_idx) = result.key_columns.iter().position(|k| k == h) {
                main_cols.push(m.key.get(key_idx).cloned().unwrap_or_default());
                old_cols.push(String::new());
            } else if let Some(change) = m.changes.iter().find(|c| c.column == *h) {
                main_cols.push(change.new_value.clone());
                old_cols.push(change.old_value.clone());
            } else {
                // Unchanged non-key column — leave empty in both
                main_cols.push(String::new());
                old_cols.push(String::new());
            }
        }

        let mut parts: Vec<String> = vec!["modified".to_string()];
        parts.extend(main_cols);
        parts.extend(old_cols);
        out.push_str(&csv_row(&parts));
        out.push('\n');
    }

    out
}

// ---------------------------------------------------------------------------
// run / main
// ---------------------------------------------------------------------------

fn run(args: Args) -> Result<()> {
    // Parse file arguments
    let (path_a, sheet_arg_a) = parse_file_arg(&args.file_a);
    let (path_b, sheet_arg_b) = parse_file_arg(&args.file_b);

    // Validate files exist
    if !path_a.exists() {
        bail!("file not found: {}", path_a.display());
    }
    if !path_b.exists() {
        bail!("file not found: {}", path_b.display());
    }

    // Validate format
    let format = args.format.to_lowercase();
    if !["text", "markdown", "json", "csv"].contains(&format.as_str()) {
        bail!(
            "unknown format '{}'. Use: text, markdown, json, csv",
            args.format
        );
    }

    // Parse skip
    let (skip_a, skip_b) = parse_skip(&args.skip)?;

    // Read file info and resolve sheets
    let info_a = metadata::read_file_info(&path_a)?;
    let info_b = metadata::read_file_info(&path_b)?;

    let sheet_a = resolve_sheet(&info_a, sheet_arg_a.as_deref())?;
    let sheet_b = resolve_sheet(&info_b, sheet_arg_b.as_deref())?;

    // Read DataFrames
    let df_a = if skip_a > 0 {
        reader::read_sheet_with_skip(&path_a, &sheet_a, skip_a)?
    } else {
        reader::read_sheet(&path_a, &sheet_a)?
    };
    let df_b = if skip_b > 0 {
        reader::read_sheet_with_skip(&path_b, &sheet_b, skip_b)?
    } else {
        reader::read_sheet(&path_b, &sheet_b)?
    };

    // Resolve key columns
    let key_columns: Vec<String> = if let Some(ref key_str) = args.key {
        let specs: Vec<String> = key_str.split(',').map(|s| s.trim().to_string()).collect();

        let df_a_cols: Vec<String> = df_a
            .get_column_names()
            .iter()
            .map(|s| s.to_string())
            .collect();
        let df_b_cols: Vec<String> = df_b
            .get_column_names()
            .iter()
            .map(|s| s.to_string())
            .collect();

        let resolved_a = filter::resolve_columns(&specs, &df_a_cols)
            .map_err(|e| anyhow::anyhow!("first file: {}", e))?;
        let resolved_b = filter::resolve_columns(&specs, &df_b_cols)
            .map_err(|e| anyhow::anyhow!("second file: {}", e))?;

        // Validate keys match between both files
        if resolved_a != resolved_b {
            bail!(
                "key columns resolve to different names: {:?} vs {:?}",
                resolved_a,
                resolved_b
            );
        }
        resolved_a
    } else {
        vec![]
    };

    // Column filtering with --cols
    let (df_a, df_b) = if let Some(ref cols_str) = args.cols {
        let specs: Vec<String> = cols_str.split(',').map(|s| s.trim().to_string()).collect();

        let df_a_cols: Vec<String> = df_a
            .get_column_names()
            .iter()
            .map(|s| s.to_string())
            .collect();
        let df_b_cols: Vec<String> = df_b
            .get_column_names()
            .iter()
            .map(|s| s.to_string())
            .collect();

        let mut selected_a = filter::resolve_columns(&specs, &df_a_cols)
            .map_err(|e| anyhow::anyhow!(e))?;
        let mut selected_b = filter::resolve_columns(&specs, &df_b_cols)
            .map_err(|e| anyhow::anyhow!(e))?;

        // Ensure key columns are included
        for key in &key_columns {
            if !selected_a.contains(key) {
                selected_a.insert(0, key.clone());
            }
            if !selected_b.contains(key) {
                selected_b.insert(0, key.clone());
            }
        }

        let df_a = df_a.select(selected_a.iter().map(|s| s.as_str()))?;
        let df_b = df_b.select(selected_b.iter().map(|s| s.as_str()))?;
        (df_a, df_b)
    } else {
        (df_a, df_b)
    };

    // Build sources and options
    let file_name_a = path_a
        .file_name()
        .map(|s| s.to_string_lossy().to_string())
        .unwrap_or_else(|| args.file_a.clone());
    let file_name_b = path_b
        .file_name()
        .map(|s| s.to_string_lossy().to_string())
        .unwrap_or_else(|| args.file_b.clone());

    let source_a = SheetSource {
        file_name: file_name_a,
        sheet_name: sheet_a,
    };
    let source_b = SheetSource {
        file_name: file_name_b,
        sheet_name: sheet_b,
    };

    let opts = DiffOptions {
        key_columns,
        tolerance: args.tolerance,
    };

    // Run diff
    let result = xlcat::diff::diff_sheets(&df_a, &df_b, &opts, source_a, source_b)?;

    // TTY detection for color
    let use_color = !args.no_color && std::io::stdout().is_terminal();

    // Format output
    let output = match format.as_str() {
        "text" => format_text(&result, use_color),
        "markdown" => format_markdown(&result),
        "json" => format_json(&result),
        "csv" => format_csv(&result),
        _ => unreachable!(),
    };

    print!("{}", output);

    // Exit 1 if differences found (diff convention)
    if result.has_differences() {
        process::exit(1);
    }

    Ok(())
}

fn main() {
    let args = Args::parse();
    if let Err(err) = run(args) {
        eprintln!("xldiff: {err}");
        process::exit(2);
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use xlcat::diff::{CellChange, DiffRow, ModifiedRow};

    // -- parse_file_arg --

    #[test]
    fn test_parse_file_arg_no_sheet() {
        let (path, sheet) = parse_file_arg("data.xlsx");
        assert_eq!(path, PathBuf::from("data.xlsx"));
        assert_eq!(sheet, None);
    }

    #[test]
    fn test_parse_file_arg_with_sheet() {
        let (path, sheet) = parse_file_arg("data.xlsx:Revenue");
        assert_eq!(path, PathBuf::from("data.xlsx"));
        assert_eq!(sheet, Some("Revenue".to_string()));
    }

    #[test]
    fn test_parse_file_arg_trailing_colon() {
        let (path, sheet) = parse_file_arg("data.xlsx:");
        assert_eq!(path, PathBuf::from("data.xlsx"));
        assert_eq!(sheet, None);
    }

    #[test]
    fn test_parse_file_arg_windows_drive() {
        let (path, sheet) = parse_file_arg("C:\\Users\\file.xlsx");
        assert_eq!(path, PathBuf::from("C:\\Users\\file.xlsx"));
        assert_eq!(sheet, None);
    }

    #[test]
    fn test_parse_file_arg_windows_drive_with_sheet() {
        let (path, sheet) = parse_file_arg("C:\\Users\\file.xlsx:Sheet2");
        assert_eq!(path, PathBuf::from("C:\\Users\\file.xlsx"));
        assert_eq!(sheet, Some("Sheet2".to_string()));
    }

    // -- parse_skip --

    #[test]
    fn test_parse_skip_single() {
        assert_eq!(parse_skip("3").unwrap(), (3, 3));
    }

    #[test]
    fn test_parse_skip_pair() {
        assert_eq!(parse_skip("3,5").unwrap(), (3, 5));
    }

    #[test]
    fn test_parse_skip_zero() {
        assert_eq!(parse_skip("0").unwrap(), (0, 0));
    }

    #[test]
    fn test_parse_skip_invalid() {
        assert!(parse_skip("abc").is_err());
    }

    // -- csv_quote --

    #[test]
    fn test_csv_quote_plain() {
        assert_eq!(csv_quote("hello"), "hello");
    }

    #[test]
    fn test_csv_quote_comma() {
        assert_eq!(csv_quote("a,b"), "\"a,b\"");
    }

    #[test]
    fn test_csv_quote_quotes() {
        assert_eq!(csv_quote("say \"hi\""), "\"say \"\"hi\"\"\"");
    }

    // -- format_text --

    #[test]
    fn test_format_text_no_diff() {
        let result = DiffResult {
            headers: vec!["a".into()],
            key_columns: vec![],
            added: vec![],
            removed: vec![],
            modified: vec![],
            source_a: SheetSource {
                file_name: "a.xlsx".into(),
                sheet_name: "Sheet1".into(),
            },
            source_b: SheetSource {
                file_name: "b.xlsx".into(),
                sheet_name: "Sheet1".into(),
            },
        };
        assert_eq!(format_text(&result, false), "No differences found.\n");
    }

    #[test]
    fn test_format_text_with_changes() {
        let result = DiffResult {
            headers: vec!["id".into(), "name".into()],
            key_columns: vec!["id".into()],
            added: vec![DiffRow {
                values: vec!["3".into(), "Charlie".into()],
            }],
            removed: vec![DiffRow {
                values: vec!["1".into(), "Alice".into()],
            }],
            modified: vec![ModifiedRow {
                key: vec!["2".into()],
                changes: vec![CellChange {
                    column: "name".into(),
                    old_value: "Bob".into(),
                    new_value: "Robert".into(),
                }],
            }],
            source_a: SheetSource {
                file_name: "a.xlsx".into(),
                sheet_name: "Sheet1".into(),
            },
            source_b: SheetSource {
                file_name: "b.xlsx".into(),
                sheet_name: "Sheet1".into(),
            },
        };
        let text = format_text(&result, false);
        assert!(text.contains("--- Sheet1 (a.xlsx)"));
        assert!(text.contains("+++ Sheet1 (b.xlsx)"));
        assert!(text.contains("Added: 1"));
        assert!(text.contains("Removed: 1"));
        assert!(text.contains("Modified: 1"));
        assert!(text.contains("- id: \"1\"  name: \"Alice\""));
        assert!(text.contains("+ id: \"3\"  name: \"Charlie\""));
        assert!(text.contains("~ [id=2]"));
        assert!(text.contains("name: \"Bob\" \u{2192} \"Robert\""));
    }

    // -- format_json --

    #[test]
    fn test_format_json_structure() {
        let result = DiffResult {
            headers: vec!["id".into(), "val".into()],
            key_columns: vec!["id".into()],
            added: vec![DiffRow {
                values: vec!["2".into(), "new".into()],
            }],
            removed: vec![],
            modified: vec![],
            source_a: SheetSource {
                file_name: "a.xlsx".into(),
                sheet_name: "S1".into(),
            },
            source_b: SheetSource {
                file_name: "b.xlsx".into(),
                sheet_name: "S1".into(),
            },
        };
        let json_str = format_json(&result);
        let parsed: Value = serde_json::from_str(&json_str).unwrap();
        assert_eq!(parsed["added"][0]["id"], "2");
        assert_eq!(parsed["added"][0]["val"], "new");
        assert!(parsed["removed"].as_array().unwrap().is_empty());
    }

    // -- format_csv --

    #[test]
    fn test_format_csv_header() {
        let result = DiffResult {
            headers: vec!["id".into(), "name".into()],
            key_columns: vec!["id".into()],
            added: vec![],
            removed: vec![],
            modified: vec![],
            source_a: SheetSource {
                file_name: "a.xlsx".into(),
                sheet_name: "S".into(),
            },
            source_b: SheetSource {
                file_name: "b.xlsx".into(),
                sheet_name: "S".into(),
            },
        };
        let csv = format_csv(&result);
        let first_line = csv.lines().next().unwrap();
        assert_eq!(first_line, "_status,id,name,_old_id,_old_name");
    }
}
