# xl-cli-tools

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

Command-line tools for viewing and editing Excel files, designed for use with LLMs and Claude Code.

Two binaries, no runtime dependencies:

- **`xlcat`** — view xlsx/xls files as markdown tables or CSV
- **`xlset`** — modify cells in existing xlsx files, preserving formatting

## xlcat demo

![xlcat demo](demo/xlcat.gif)

## xlset demo

![xlset demo](demo/xlset.gif)

## Install

### Pre-built binaries (macOS)

Download from [Releases](https://github.com/LouLouLibs/xl-cli-tools/releases):

```bash
# Apple Silicon
curl -L https://github.com/LouLouLibs/xl-cli-tools/releases/latest/download/xlcat-aarch64-apple-darwin -o ~/.local/bin/xlcat
curl -L https://github.com/LouLouLibs/xl-cli-tools/releases/latest/download/xlset-aarch64-apple-darwin -o ~/.local/bin/xlset
chmod +x ~/.local/bin/xlcat ~/.local/bin/xlset

# Intel Mac
curl -L https://github.com/LouLouLibs/xl-cli-tools/releases/latest/download/xlcat-x86_64-apple-darwin -o ~/.local/bin/xlcat
curl -L https://github.com/LouLouLibs/xl-cli-tools/releases/latest/download/xlset-x86_64-apple-darwin -o ~/.local/bin/xlset
chmod +x ~/.local/bin/xlcat ~/.local/bin/xlset
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

Both tools include Claude Code skills (`/xls` and `/xlset`) for seamless use in conversations. Claude can view spreadsheets, analyze data, and make targeted edits.

## Exit codes

| Code | xlcat | xlset |
|------|-------|-------|
| 0 | Success | Success |
| 1 | Runtime error | Runtime error |
| 2 | Invalid arguments | Invalid arguments |

## Tech

- **xlcat:** calamine (Excel reading) + polars (DataFrames, statistics) + clap
- **xlset:** umya-spreadsheet (round-trip Excel editing) + clap
- Both compile to single static binaries with no runtime dependencies
