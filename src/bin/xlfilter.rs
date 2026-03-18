use std::path::PathBuf;
use std::process;

use anyhow::Result;
use clap::Parser;

use xlcat::filter::{
    parse_filter_expr, parse_sort_spec, filter_pipeline, FilterOptions,
};
use xlcat::formatter;
use xlcat::metadata;
use xlcat::metadata::SheetInfo;
use xlcat::reader;

#[derive(Parser)]
#[command(
    name = "xlfilter",
    about = "Filter and query Excel spreadsheet data",
    version
)]
struct Args {
    /// Path to .xls or .xlsx file
    file: PathBuf,

    /// Select columns by letter (A,B,D) or header name (State,Amount)
    #[arg(long)]
    cols: Option<String>,

    /// Filter rows. Multiple = AND. Operators: = != > < >= <= ~ (contains, case-insensitive) !~ (not contains)
    #[arg(long = "where")]
    filters: Vec<String>,

    /// Sort by column. Format: col:dir (dir = asc or desc, default asc)
    #[arg(long)]
    sort: Option<String>,

    /// Max rows in output (applied after filtering)
    #[arg(long)]
    limit: Option<usize>,

    /// First N rows (applied before filtering)
    #[arg(long)]
    head: Option<usize>,

    /// Last N rows (applied before filtering)
    #[arg(long)]
    tail: Option<usize>,

    /// Target sheet by name or 0-based index (default: first sheet)
    #[arg(long)]
    sheet: Option<String>,

    /// Output as CSV instead of markdown table
    #[arg(long)]
    csv: bool,

    /// Skip first N data rows (for metadata/title rows above the real header)
    #[arg(long)]
    skip: Option<usize>,
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

fn run(args: Args) -> Result<()> {
    // Validate file exists
    if !args.file.exists() {
        anyhow::bail!(ArgError(format!(
            "file not found: {}",
            args.file.display()
        )));
    }

    // Validate mutually exclusive flags
    if args.head.is_some() && args.tail.is_some() {
        anyhow::bail!(ArgError(
            "--head and --tail are mutually exclusive".to_string()
        ));
    }

    // Read file metadata to resolve sheet name
    let info = metadata::read_file_info(&args.file)?;
    let sheet_name = resolve_sheet(&info, args.sheet.as_deref())?;

    // Read sheet into DataFrame
    let df = reader::read_sheet(&args.file, &sheet_name)?;

    if df.height() == 0 {
        eprintln!("0 rows");
        let sheet_info = info
            .sheets
            .iter()
            .find(|s| s.name == sheet_name)
            .cloned()
            .unwrap_or(SheetInfo {
                name: sheet_name,
                rows: 0,
                cols: 0,
            });
        println!("{}", formatter::format_empty_sheet(&sheet_info));
        return Ok(());
    }

    // Parse filter expressions
    let filters: Vec<_> = args
        .filters
        .iter()
        .map(|s| parse_filter_expr(s))
        .collect::<Result<Vec<_>, _>>()
        .map_err(|e| anyhow::anyhow!(ArgError(e)))?;

    // Parse sort spec
    let sort = args
        .sort
        .as_deref()
        .map(parse_sort_spec)
        .transpose()
        .map_err(|e| anyhow::anyhow!(ArgError(e)))?;

    // Parse column selection
    let cols = args.cols.map(|s| {
        s.split(',')
            .map(|c| c.trim().to_string())
            .collect::<Vec<_>>()
    });

    // Run pipeline
    let opts = FilterOptions {
        filters,
        cols,
        sort,
        limit: args.limit,
        head: args.head,
        tail: args.tail,
    };
    let result = filter_pipeline(df, &opts)?;

    // Output row count to stderr
    eprintln!("{} rows", result.height());

    // Format output
    if result.height() == 0 {
        println!("{}", formatter::format_data_table(&result));
    } else if args.csv {
        print!("{}", formatter::format_csv(&result));
    } else {
        println!("{}", formatter::format_data_table(&result));
    }

    Ok(())
}

/// Resolve sheet name from --sheet flag or default to first sheet.
fn resolve_sheet(info: &metadata::FileInfo, sheet_arg: Option<&str>) -> Result<String> {
    match sheet_arg {
        None => {
            info.sheets
                .first()
                .map(|s| s.name.clone())
                .ok_or_else(|| anyhow::anyhow!("workbook has no sheets"))
        }
        Some(s) => {
            if let Some(sheet) = info.sheets.iter().find(|si| si.name == s) {
                return Ok(sheet.name.clone());
            }
            if let Ok(idx) = s.parse::<usize>() {
                if let Some(sheet) = info.sheets.get(idx) {
                    return Ok(sheet.name.clone());
                }
            }
            let names: Vec<_> = info.sheets.iter().map(|s| s.name.as_str()).collect();
            anyhow::bail!(ArgError(format!(
                "sheet '{}' not found. Available sheets: {}",
                s,
                names.join(", ")
            )))
        }
    }
}

fn main() {
    let args = Args::parse();
    if let Err(err) = run(args) {
        if err.downcast_ref::<ArgError>().is_some() {
            eprintln!("xlfilter: {err}");
            process::exit(2);
        }
        eprintln!("xlfilter: {err}");
        process::exit(1);
    }
}
