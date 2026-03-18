<div align="center">

<h1>xl-cli-tools</h1>
<h3>View and edit Excel files from the command line</h3>

[![Vibecoded](https://img.shields.io/badge/vibecoded-%E2%9C%A8-blueviolet)](https://claude.ai)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

<table>
<tr>
<td align="center" width="33%"><strong>xlcat</strong> — view</td>
<td align="center" width="33%"><strong>xlset</strong> — edit</td>
<td align="center" width="33%"><strong>xlfilter</strong> — query</td>
</tr>
<tr>
<td><img src="demo/xlcat.gif" alt="xlcat demo" /></td>
<td><img src="demo/xlset.gif" alt="xlset demo" /></td>
<td><img src="demo/xlfilter.gif" alt="xlfilter demo" /></td>
</tr>
</table>

</div>

***

Three binaries, no runtime dependencies:

- **`xlcat`** — view xlsx/xls files as markdown tables or CSV
- **`xlset`** — modify cells in existing xlsx files, preserving formatting
- **`xlfilter`** — filter, sort, and query rows from spreadsheets

## Installation

### Pre-built binaries (macOS)

Download from [Releases](https://github.com/LouLouLibs/xl-cli-tools/releases):

```bash
# Apple Silicon
curl -L https://github.com/LouLouLibs/xl-cli-tools/releases/latest/download/xlcat-aarch64-apple-darwin -o ~/.local/bin/xlcat
curl -L https://github.com/LouLouLibs/xl-cli-tools/releases/latest/download/xlset-aarch64-apple-darwin -o ~/.local/bin/xlset
curl -L https://github.com/LouLouLibs/xl-cli-tools/releases/latest/download/xlfilter-aarch64-apple-darwin -o ~/.local/bin/xlfilter
chmod +x ~/.local/bin/xlcat ~/.local/bin/xlset ~/.local/bin/xlfilter

# Intel Mac
curl -L https://github.com/LouLouLibs/xl-cli-tools/releases/latest/download/xlcat-x86_64-apple-darwin -o ~/.local/bin/xlcat
curl -L https://github.com/LouLouLibs/xl-cli-tools/releases/latest/download/xlset-x86_64-apple-darwin -o ~/.local/bin/xlset
curl -L https://github.com/LouLouLibs/xl-cli-tools/releases/latest/download/xlfilter-x86_64-apple-darwin -o ~/.local/bin/xlfilter
chmod +x ~/.local/bin/xlcat ~/.local/bin/xlset ~/.local/bin/xlfilter
```

### From source

```bash
cargo install --path .
```

Requires Rust 1.85+.

## xlcat — View Excel Files

```bash
# Overview: metadata, schema, first/last 25 rows
xlcat report.xlsx

# Column names and types only
xlcat report.xlsx --schema

# Summary statistics (count, mean, std, min, max, median)
xlcat report.xlsx --describe

# Pick a sheet in a multi-sheet workbook
xlcat report.xlsx --sheet Revenue

# First 10 rows
xlcat report.xlsx --head 10

# Last 5 rows
xlcat report.xlsx --tail 5

# Both
xlcat report.xlsx --head 10 --tail 5

# All rows (overrides large-file gate)
xlcat report.xlsx --all

# CSV output for piping
xlcat report.xlsx --csv
xlcat report.xlsx --csv --head 100 > subset.csv
```

### Example output

```
# File: sales.xlsx (245 KB)
# Sheets: 1

## Sheet: Q1 (1240 rows x 4 cols)

| Column  | Type   |
|---------|--------|
| date    | Date   |
| region  | String |
| amount  | Float  |
| units   | Int    |

| date       | region | amount  | units |
|------------|--------|---------|-------|
| 2024-01-01 | East   | 1234.56 | 100   |
| 2024-01-02 | West   | 987.00  | 75    |
... (1190 rows omitted) ...
| 2024-12-30 | East   | 1100.00 | 92    |
| 2024-12-31 | West   | 1250.75 | 110   |
```

### Adaptive defaults

- **Single sheet, <=50 rows:** shows all data
- **Single sheet, >50 rows:** first 25 + last 25 rows
- **Multiple sheets:** lists schemas, pick one with `--sheet`
- **Large file (>1MB):** schema + first 25 rows (override with `--max-size 5M`)

## xlfilter — Query and Filter

```bash
# Filter rows by value
xlfilter data.xlsx --where "State=CA"

# Numeric comparisons
xlfilter data.xlsx --where "Amount>1000"

# Multiple filters (AND)
xlfilter data.xlsx --where "State=CA" --where "Amount>1000"

# Select columns (by name or letter)
xlfilter data.xlsx --cols State,Amount,Year
xlfilter data.xlsx --cols A,C,D

# Sort results
xlfilter data.xlsx --sort "Amount:desc"

# Limit output
xlfilter data.xlsx --sort "Amount:desc" --limit 10

# Contains filter (case-insensitive)
xlfilter data.xlsx --where "Name~john"

# Head/tail (applied before filtering)
xlfilter data.xlsx --head 100 --where "Status=Active"

# Skip metadata rows above the real header
xlfilter data.xlsx --skip 2

# CSV output for piping
xlfilter data.xlsx --where "Status!=Draft" --csv | other-tool

# Target a specific sheet
xlfilter data.xlsx --sheet Revenue --where "Amount>5000"
```

### Filter operators

| Operator | Meaning | Example |
|----------|---------|---------|
| `=` | Equals | `State=CA` |
| `!=` | Not equals | `Status!=Draft` |
| `>` | Greater than | `Amount>1000` |
| `<` | Less than | `Year<2024` |
| `>=` | Greater or equal | `Score>=90` |
| `<=` | Less or equal | `Price<=50` |
| `~` | Contains (case-insensitive) | `Name~john` |
| `!~` | Not contains | `Name!~test` |

Numeric columns compare numerically; string columns compare lexicographically. Row count is printed to stderr.

## xlset — Edit Excel Cells

```bash
# Set a single cell
xlset report.xlsx A2=42

# Set multiple cells
xlset report.xlsx A2=42 B2="hello world" C2=true

# Preserve leading zeros with type tag
xlset report.xlsx A2:str=07401

# Target a specific sheet
xlset report.xlsx --sheet Revenue A2=42

# Write to a new file (don't modify original)
xlset report.xlsx --output modified.xlsx A2=42

# Bulk update from CSV
xlset report.xlsx --from updates.csv

# Bulk from stdin
echo "A1,42" | xlset report.xlsx --from -
```

### Type inference

Values are auto-detected: `42` becomes a number, `true` becomes boolean, `2024-01-15` becomes a date. Override with tags when needed:

| Tag | Effect |
|-----|--------|
| `:str` | Force string (`A1:str=07401` preserves leading zero) |
| `:num` | Force number |
| `:bool` | Force boolean |
| `:date` | Force date |

### CSV format for `--from`

```csv
cell,value
A1,42
B2,hello world
C3:str,07401
D4,"value with, comma"
```

### What gets preserved

xlset modifies only the cells you specify. Everything else is untouched: formatting, formulas, charts, conditional formatting, data validation, merged cells, images.

## Claude Code integration

Claude Code skills (`/xlcat` and `/xlset`) are available in [claude-skills](https://github.com/LouLouLibs/claude-skills). Claude can view spreadsheets, analyze data, and make targeted edits in conversations.

## Exit codes

| Code | Meaning |
|------|---------|
| 0 | Success |
| 1 | Runtime error |
| 2 | Invalid arguments |

## License

MIT
