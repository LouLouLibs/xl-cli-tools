mod formatter;
mod metadata;
mod reader;

use anyhow::Result;
use clap::Parser;
use polars::prelude::*;
use std::path::PathBuf;
use std::process;

use metadata::{FileInfo, SheetInfo};

#[derive(Parser, Debug)]
#[command(name = "xlcat", about = "View Excel files in the terminal")]
struct Cli {
    /// Path to .xls or .xlsx file
    file: PathBuf,

    /// Show only column names and types
    #[arg(long)]
    schema: bool,

    /// Show summary statistics
    #[arg(long)]
    describe: bool,

    /// Show first N rows
    #[arg(long)]
    head: Option<usize>,

    /// Show last N rows
    #[arg(long)]
    tail: Option<usize>,

    /// Show all rows (overrides large-file gate)
    #[arg(long)]
    all: bool,

    /// Select sheet by name or 0-based index
    #[arg(long)]
    sheet: Option<String>,

    /// Large-file threshold (default: 1M). Accepts: 500K, 1M, 10M, 1G
    #[arg(long, default_value = "1M", value_parser = parse_size)]
    max_size: u64,

    /// Output as CSV instead of markdown
    #[arg(long)]
    csv: bool,
}

fn parse_size(s: &str) -> Result<u64, String> {
    let s = s.trim();
    let (num_part, multiplier) = if s.ends_with('G') || s.ends_with('g') {
        (&s[..s.len() - 1], 1_073_741_824u64)
    } else if s.ends_with("GB") || s.ends_with("gb") {
        (&s[..s.len() - 2], 1_073_741_824u64)
    } else if s.ends_with('M') || s.ends_with('m') {
        (&s[..s.len() - 1], 1_048_576u64)
    } else if s.ends_with("MB") || s.ends_with("mb") {
        (&s[..s.len() - 2], 1_048_576u64)
    } else if s.ends_with('K') || s.ends_with('k') {
        (&s[..s.len() - 1], 1_024u64)
    } else if s.ends_with("KB") || s.ends_with("kb") {
        (&s[..s.len() - 2], 1_024u64)
    } else {
        (s, 1u64)
    };
    let num: f64 = num_part.parse().map_err(|_| format!("Invalid size: {s}"))?;
    Ok((num * multiplier as f64) as u64)
}

// ---------------------------------------------------------------------------
// ArgError — used for user-facing flag/argument errors (exit code 2)
// ---------------------------------------------------------------------------

#[derive(Debug)]
struct ArgError(String);

impl std::fmt::Display for ArgError {
    fn fmt(&self, f: &mut std::fmt::Formatter) -> std::fmt::Result {
        write!(f, "{}", self.0)
    }
}

impl std::error::Error for ArgError {}

// ---------------------------------------------------------------------------
// Sheet resolution
// ---------------------------------------------------------------------------

enum SheetTarget {
    Single(usize),
    ListAll,
}

// ---------------------------------------------------------------------------
// run() — main orchestration
// ---------------------------------------------------------------------------

fn run(cli: &Cli) -> Result<()> {
    // 1. Validate flag combinations
    if cli.schema && cli.describe {
        return Err(ArgError("--schema and --describe are mutually exclusive".into()).into());
    }
    if (cli.schema || cli.describe)
        && (cli.head.is_some() || cli.tail.is_some() || cli.all)
    {
        return Err(ArgError(
            "--schema/--describe cannot be combined with --head, --tail, or --all".into(),
        )
        .into());
    }
    if (cli.schema || cli.describe) && cli.csv {
        return Err(ArgError(
            "--csv cannot be combined with --schema or --describe".into(),
        )
        .into());
    }

    // 2. Read file metadata
    let info = metadata::read_file_info(&cli.file)?;
    let file_name = cli
        .file
        .file_name()
        .map(|s| s.to_string_lossy().to_string())
        .unwrap_or_else(|| cli.file.display().to_string());

    // 3. Resolve sheet target
    let target = resolve_sheet_target(cli, &info)?;

    match target {
        SheetTarget::Single(idx) => {
            let sheet = &info.sheets[idx];
            let df = reader::read_sheet(&cli.file, &sheet.name)?;
            render_single_sheet(cli, &file_name, &info, sheet, &df)?;
        }
        SheetTarget::ListAll => {
            if cli.describe {
                // --describe on multi-sheet: iterate all sheets
                let mut out = formatter::format_header(&file_name, &info);
                out.push('\n');
                for sheet in &info.sheets {
                    let df = reader::read_sheet(&cli.file, &sheet.name)?;
                    if sheet.rows == 0 && sheet.cols == 0 {
                        out.push_str(&formatter::format_empty_sheet(sheet));
                    } else {
                        out.push_str(&formatter::format_schema(sheet, &df));
                        out.push_str(&formatter::format_describe(&df));
                    }
                    out.push('\n');
                }
                print!("{out}");
            } else {
                // Default multi-sheet: list schemas
                let mut pairs: Vec<(&SheetInfo, DataFrame)> = Vec::new();
                for sheet in &info.sheets {
                    let df = reader::read_sheet(&cli.file, &sheet.name)?;
                    pairs.push((sheet, df));
                }
                let out = formatter::format_sheet_listing(&file_name, &info, &pairs);
                print!("{out}");
            }
        }
    }

    Ok(())
}

fn resolve_sheet_target(cli: &Cli, info: &FileInfo) -> Result<SheetTarget> {
    if let Some(ref sheet_arg) = cli.sheet {
        // Try name match first
        if let Some(idx) = info.sheets.iter().position(|s| s.name == *sheet_arg) {
            return Ok(SheetTarget::Single(idx));
        }
        // Try 0-based index
        if let Ok(idx) = sheet_arg.parse::<usize>() {
            if idx < info.sheets.len() {
                return Ok(SheetTarget::Single(idx));
            }
            return Err(ArgError(format!(
                "Sheet index {idx} out of range (file has {} sheets)",
                info.sheets.len()
            ))
            .into());
        }
        return Err(ArgError(format!("Sheet not found: {sheet_arg}")).into());
    }

    if info.sheets.len() == 1 {
        return Ok(SheetTarget::Single(0));
    }

    // Multi-sheet, no --sheet specified
    let has_row_flags = cli.all || cli.head.is_some() || cli.tail.is_some() || cli.csv;
    if has_row_flags {
        return Err(ArgError(
            "Multiple sheets found. Use --sheet <name> to select one before using --all, --head, --tail, or --csv.".into(),
        )
        .into());
    }

    Ok(SheetTarget::ListAll)
}

fn render_single_sheet(
    cli: &Cli,
    file_name: &str,
    info: &FileInfo,
    sheet: &SheetInfo,
    df: &DataFrame,
) -> Result<()> {
    // CSV mode: apply row selection, output CSV, done
    if cli.csv {
        let selected = apply_row_selection(cli, info, df);
        let csv_out = formatter::format_csv(&selected);
        print!("{csv_out}");
        return Ok(());
    }

    let mut out = formatter::format_header(file_name, info);
    out.push('\n');

    // Completely empty sheet (0 rows, 0 cols)
    if sheet.rows == 0 && sheet.cols == 0 {
        out.push_str(&formatter::format_empty_sheet(sheet));
        print!("{out}");
        return Ok(());
    }

    // Header-only sheet (has columns but 0 data rows)
    if df.height() == 0 {
        out.push_str(&formatter::format_schema(sheet, df));
        out.push_str("\n(no data rows)\n");
        print!("{out}");
        return Ok(());
    }

    if cli.schema {
        out.push_str(&formatter::format_schema(sheet, df));
    } else if cli.describe {
        out.push_str(&formatter::format_schema(sheet, df));
        out.push_str(&formatter::format_describe(df));
    } else {
        // Data mode
        out.push_str(&formatter::format_schema(sheet, df));
        out.push('\n');
        out.push_str(&format_data_with_selection(cli, info, df));
    }

    print!("{out}");
    Ok(())
}

/// Format data output with row selection logic.
fn format_data_with_selection(cli: &Cli, info: &FileInfo, df: &DataFrame) -> String {
    let total = df.height();

    // --all: show everything
    if cli.all {
        return formatter::format_data_table(df);
    }

    // Explicit --head and/or --tail
    if cli.head.is_some() || cli.tail.is_some() {
        let head_n = cli.head.unwrap_or(0);
        let tail_n = cli.tail.unwrap_or(0);
        if head_n + tail_n >= total || (head_n == 0 && tail_n == 0) {
            return formatter::format_data_table(df);
        }
        // If only --head, show first N
        if cli.tail.is_none() {
            let head_df = df.head(Some(head_n));
            return formatter::format_data_table(&head_df);
        }
        // If only --tail, show last N
        if cli.head.is_none() {
            let tail_df = df.tail(Some(tail_n));
            return formatter::format_data_table(&tail_df);
        }
        // Both specified
        return formatter::format_head_tail(df, head_n, tail_n);
    }

    // Large file gate: file_size > max_size and no explicit flags
    if info.file_size > cli.max_size {
        let mut out = formatter::format_head_tail(df, 25, 0);
        out.push_str(&format!(
            "\nLarge file ({}) — showing first 25 of {total} rows. Use --all to see everything.\n",
            metadata::format_file_size(info.file_size)
        ));
        return out;
    }

    // Adaptive default: <=50 rows show all, >50 show head 25 + tail 25
    if total <= 50 {
        formatter::format_data_table(df)
    } else {
        formatter::format_head_tail(df, 25, 25)
    }
}

/// Apply row selection for CSV mode — returns a (possibly sliced) DataFrame.
fn apply_row_selection(cli: &Cli, info: &FileInfo, df: &DataFrame) -> DataFrame {
    let total = df.height();

    if cli.all {
        return df.clone();
    }

    if cli.head.is_some() || cli.tail.is_some() {
        let head_n = cli.head.unwrap_or(0);
        let tail_n = cli.tail.unwrap_or(0);

        if head_n + tail_n >= total || (head_n == 0 && tail_n == 0) {
            return df.clone();
        }

        if cli.tail.is_none() {
            return df.head(Some(head_n));
        }
        if cli.head.is_none() {
            return df.tail(Some(tail_n));
        }

        // Both head and tail: combine
        let head_df = df.head(Some(head_n));
        let tail_df = df.tail(Some(tail_n));
        return head_df.vstack(&tail_df).unwrap_or_else(|_| df.clone());
    }

    // Large file gate
    if info.file_size > cli.max_size {
        return df.head(Some(25));
    }

    // Adaptive default
    if total <= 50 {
        df.clone()
    } else {
        let head_df = df.head(Some(25));
        let tail_df = df.tail(Some(25));
        head_df.vstack(&tail_df).unwrap_or_else(|_| df.clone())
    }
}

// ---------------------------------------------------------------------------
// main()
// ---------------------------------------------------------------------------

fn main() {
    let cli = Cli::parse();
    if let Err(err) = run(&cli) {
        // Check if the root cause is an ArgError
        if err.downcast_ref::<ArgError>().is_some() {
            eprintln!("xlcat: {err}");
            process::exit(2);
        }
        eprintln!("xlcat: {err}");
        process::exit(1);
    }
}
