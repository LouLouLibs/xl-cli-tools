<div align="center">

<h1>xl-cli-tools</h1>
<h3>View, edit, query, and diff Excel files from the command line</h3>

[![Vibecoded](https://img.shields.io/badge/vibecoded-%E2%9C%A8-blueviolet)](https://claude.ai)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

<table>
<tr>
<td align="center" width="50%"><strong>xlcat</strong> — view</td>
<td align="center" width="50%"><strong>xlset</strong> — edit</td>
</tr>
<tr>
<td><img src="demo/xlcat.gif" alt="xlcat demo" /></td>
<td><img src="demo/xlset.gif" alt="xlset demo" /></td>
</tr>
<tr>
<td align="center" width="50%"><strong>xlfilter</strong> — query</td>
<td align="center" width="50%"><strong>xldiff</strong> — compare</td>
</tr>
<tr>
<td><img src="demo/xlfilter.gif" alt="xlfilter demo" /></td>
<td><img src="demo/xldiff.gif" alt="xldiff demo" /></td>
</tr>
</table>

</div>

***

[**xlcat**](#xlcat--view-excel-files) · [**xlset**](#xlset--edit-excel-cells) · [**xlfilter**](#xlfilter--query-and-filter) · [**xldiff**](#xldiff--compare-two-sheets) · [**Install**](#installation) · [**Claude Code**](#claude-code-integration)

***

Four binaries, no runtime dependencies:

```bash
# View a spreadsheet
xlcat report.xlsx

# Edit a cell
xlset report.xlsx B3=42

# Filter rows
xlfilter data.xlsx --where "Amount>1000" --sort "Amount:desc"

# Diff two files
xldiff old.xlsx new.xlsx --key ID
```

## Installation

### Pre-built binaries (macOS)

Download from [Releases](https://github.com/LouLouLibs/xl-cli-tools/releases):

```bash
# Apple Silicon
for tool in xlcat xlset xlfilter xldiff; do
  curl -L "https://github.com/LouLouLibs/xl-cli-tools/releases/latest/download/${tool}-aarch64-apple-darwin" \
    -o ~/.local/bin/$tool
done
chmod +x ~/.local/bin/xl{cat,set,filter,diff}

# Intel Mac — replace aarch64 with x86_64
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

# First 10 rows / last 5 rows / both
xlcat report.xlsx --head 10
xlcat report.xlsx --tail 5
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

## xldiff — Compare Two Sheets

```bash
# Positional diff (whole-row comparison)
xldiff old.xlsx new.xlsx

# Key-based diff (match rows by ID, compare cell by cell)
xldiff old.xlsx new.xlsx --key ID

# Composite key
xldiff old.xlsx new.xlsx --key Date,Ticker

# Compare sheets within the same file
xldiff report.xlsx:Q1 report.xlsx:Q2

# Float tolerance (differences <= 0.01 treated as equal)
xldiff old.xlsx new.xlsx --key ID --tolerance 0.01

# Only compare specific columns
xldiff old.xlsx new.xlsx --key ID --cols Name,Salary

# Skip metadata rows (different skip per file)
xldiff file1.xlsx file2.xlsx --skip 3,5

# Output formats
xldiff old.xlsx new.xlsx --key ID --format markdown
xldiff old.xlsx new.xlsx --key ID --format json
xldiff old.xlsx new.xlsx --key ID --format csv
```

### Example output

```
--- Sheet1 (old.xlsx)
+++ Sheet1 (new.xlsx)

Added: 1 | Removed: 1 | Modified: 2

- ID: "3"  Name: "Charlie"  Department: "Engineering"  Salary: "88000"
+ ID: "5"  Name: "Eve"  Department: "Marketing"  Salary: "70000"
~ ID: "1"
    Salary: "95000" → "98000"
~ ID: "2"
    Department: "Marketing" → "Design"
    Salary: "72000" → "75000"
```

### Diff modes

**Positional (no `--key`):** Every column defines row identity. Reports added/removed rows only.

**Key-based (`--key`):** Match rows by key columns, compare remaining columns cell by cell. Reports added, removed, and modified rows with per-cell changes. Supports composite keys, duplicate key detection, and float tolerance.

### Exit codes (diff convention)

| Code | Meaning |
|------|---------|
| 0 | No differences |
| 1 | Differences found |
| 2 | Error |

## Claude Code integration

Claude Code skills are available in [claude-skills](https://github.com/LouLouLibs/claude-skills). Claude can view spreadsheets, analyze data, filter rows, compare files, and make targeted edits in conversations.

## Exit codes

| Code | Meaning |
|------|---------|
| 0 | Success |
| 1 | Runtime error (xldiff: differences found) |
| 2 | Invalid arguments |

## License

MIT
