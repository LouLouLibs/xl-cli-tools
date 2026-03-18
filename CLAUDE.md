# xl-cli-tool

Rust CLI tools for working with Excel (.xlsx/.xls) files, designed for LLM-assisted analysis (Claude Code).

## Current tools
- **xlcat** — view/inspect spreadsheets (metadata, schema, stats, data as markdown tables or CSV)
- **xlset** — write cells preserving formatting/formulas/structure

## Architecture
- Shared library in `src/lib.rs` (reader, formatter, metadata, cell, writer modules)
- Binary entry points in `src/bin/`
- Dependencies: calamine (read), umya-spreadsheet (write), polars (data), clap (CLI)

## Related tools (separate repos)
- **xldiff** (Go, `/Users/loulou/Dropbox/projects_claude/go-xldiff`) — compare two sheets. Future: port to Rust in this repo.

## Planned
- **xlfilter** — query/filter rows and columns from spreadsheets. See `docs/xlfilter-design.md`.

## Conventions
- Exit codes: 0 = success, 1 = runtime error, 2 = invalid arguments
- Output: markdown tables by default, CSV with `--csv`
- Rust edition 2024
