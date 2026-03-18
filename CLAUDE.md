# xl-cli-tool

Rust CLI tools for working with Excel (.xlsx/.xls) files, designed for LLM-assisted analysis (Claude Code).

## Current tools
- **xlcat** — view/inspect spreadsheets (metadata, schema, stats, data as markdown tables or CSV)
- **xlset** — write cells preserving formatting/formulas/structure
- **xlfilter** — query/filter rows and columns from spreadsheets
- **xldiff** — compare two sheets, report added/removed/modified rows

## Architecture
- Shared library in `src/lib.rs` (reader, formatter, metadata, cell, writer, filter, diff modules)
- Binary entry points in `src/bin/`
- Dependencies: calamine (read), umya-spreadsheet (write), polars (data), clap (CLI), serde_json (JSON output)

## Conventions
- Exit codes: 0 = success, 1 = runtime error, 2 = invalid arguments
  - Exception: xldiff uses diff(1) convention: 0 = no differences, 1 = differences found, 2 = error
- Output: markdown tables by default, CSV with `--csv`
- Rust edition 2024
