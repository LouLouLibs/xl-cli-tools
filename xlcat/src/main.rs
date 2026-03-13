mod formatter;
mod metadata;
mod reader;

use anyhow::Result;
use clap::Parser;
use std::path::PathBuf;

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

fn main() -> Result<()> {
    let cli = Cli::parse();

    // Validate flag combinations
    let mode_count = cli.schema as u8 + cli.describe as u8;
    if mode_count > 1 {
        anyhow::bail!("--schema and --describe are mutually exclusive");
    }
    if (cli.schema || cli.describe) && (cli.head.is_some() || cli.tail.is_some() || cli.all) {
        anyhow::bail!("--schema and --describe cannot be combined with --head, --tail, or --all");
    }
    if (cli.schema || cli.describe) && cli.csv {
        anyhow::bail!("--csv can only be used in data mode (not with --schema or --describe)");
    }

    eprintln!("xlcat: not yet implemented");
    std::process::exit(1);
}
