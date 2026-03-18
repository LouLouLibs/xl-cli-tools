# xldiff — Excel Sheet Comparison Tool

Port of [go-xldiff](../../../go-xldiff) to Rust, leveraging polars for type-aware comparison.

## CLI Interface

```
xldiff <FILE1>[:<SHEET>] <FILE2>[:<SHEET>] [flags]
```

Two positional file arguments, each optionally suffixed with `:SheetName` or `:0` (0-based index). The same file may appear twice with different sheets (e.g., `report.xlsx:Q1 report.xlsx:Q2`).

Sheet selection uses `rfind(':')` to split on the last colon, so Windows drive-letter paths (e.g., `C:\file.xlsx`) are handled correctly. The suffix is tried as a sheet name first, then as a 0-based numeric index.

### Flags

| Flag | Type | Default | Description |
|------|------|---------|-------------|
| `--key <COLS>` | string | *(none)* | Comma-separated key columns (header name or column letter). Omit for positional mode. |
| `--cols <COLS>` | string | *(all)* | Columns to compare (header name or column letter). Key columns always included. |
| `--skip <N>[,<M>]` | string | `0` | Rows to skip before header. Single value applies to both files; comma pair sets each independently. |
| `--no-header` | bool | false | Treat first row as data; generate synthetic `column_0`, `column_1`, … headers (matching reader.rs convention). |
| `--tolerance <F>` | f64 | *(none)* | Numeric tolerance for float comparisons (e.g., `0.01`). |
| `--format <FMT>` | string | `text` | Output format: `text`, `markdown`, `json`, `csv`. |
| `--no-color` | bool | false | Force-disable ANSI colors (auto-detected by default). |

**Column resolution priority** (consistent with `filter.rs::resolve_column()`):
1. Exact header name match
2. Case-insensitive header name match
3. Column letter (A, B, AA) if purely alphabetic

This means `--key A` matches a header named "A" before column position 0. To select the first column by position, use `--cols` with letter "A" only when no header is named "A".

### Exit Codes

xldiff follows the `diff(1)` convention, which differs from xlcat/xlset/xlfilter:

- **0** — no differences
- **1** — differences found
- **2** — error (invalid arguments or runtime failure)

The other tools in this repo use exit 1 for runtime errors and exit 2 for argument errors. xldiff merges both error types into exit 2 because exit 1 is reserved for "differences found" — the standard convention for diff tools, enabling `if ! xldiff ...; then ...` in scripts.

## Data Structures

```rust
pub struct DiffResult {
    pub headers: Vec<String>,       // Columns that were compared (filtered set if --cols used)
    pub key_columns: Vec<String>,
    pub added: Vec<DiffRow>,
    pub removed: Vec<DiffRow>,
    pub modified: Vec<ModifiedRow>,
    pub source_a: SheetSource,
    pub source_b: SheetSource,
}

pub struct SheetSource {
    pub file_name: String,
    pub sheet_name: String,
}

pub struct DiffRow {
    pub values: Vec<String>,        // Cell values as display strings
}

pub struct ModifiedRow {
    pub key: Vec<String>,
    pub changes: Vec<CellChange>,
}

pub struct CellChange {
    pub column: String,
    pub old_value: String,
    pub new_value: String,
}
```

`DiffResult` is format-agnostic. Each output formatter consumes it independently. `SheetSource` carries file/sheet metadata for headers.

Values are converted to display strings during diff construction. Type-aware comparison (tolerance, numeric equality) happens during the diff phase; the result stores only the string representations.

`headers` reflects the effective column set after `--cols` filtering, so output formatters use it directly without further filtering.

## Diff Algorithm

### Mismatched Headers

Before diffing, the tool checks whether both DataFrames share the same column names (after `--cols` filtering). Behavior:

- **Key-based mode:** key columns must exist in both files or the tool exits with an error. Non-key columns are compared by name — columns present in only one file are reported as added/removed columns in a warning on stderr, and comparison proceeds on the intersection. `DiffResult.headers` contains the union.
- **Positional mode:** columns are matched by position. If column counts differ, the tool warns on stderr and pads the shorter side with empty values.

### Positional Mode (no `--key`)

1. Convert each row to a string: join all compared columns with a null-byte separator.
2. Build frequency maps (`HashMap<RowHash, usize>`) for both tables.
3. Walk table A: if B has remaining copies, consume one; otherwise mark as removed.
4. Walk table B: if A has remaining copies, consume one; otherwise mark as added.
5. No "modified" rows — any cell change produces a different hash.

### Key-Based Mode (`--key` provided)

The implementation uses HashMap-based row matching (like the Go version) rather than polars joins. This avoids complications with join column type mismatches, row reordering, and duplicate-key cartesian products. Polars is used for reading and column selection, not for the join itself.

1. Resolve key column names/letters to column indices.
2. Build `HashMap<Vec<String>, RowData>` for both tables, keyed by the concatenated key column values.
3. For each key in A not in B → removed.
4. For each key in B not in A → added.
5. For each key in both → compare non-key columns cell by cell:
   - **Numeric (Float64):** if `--tolerance` is set, equal when `|a - b| <= tolerance`; otherwise exact. NaN == NaN is treated as equal; NaN vs non-NaN is a difference.
   - **Integer vs Float cross-type:** if one file infers a column as Int64 and the other as Float64, both values are compared as f64. Tolerance applies if set.
   - **Integer (Int64):** exact comparison (tolerance does not apply to pure integer columns).
   - **String:** exact string comparison.
   - **Datetime:** compare as timestamps.
   - **Boolean:** exact match.
   - **Null vs value:** treated as different. Null vs null: equal.
6. Rows with any cell difference become a `ModifiedRow` with a list of `CellChange`s.
7. Duplicate keys produce a warning on stderr. Last occurrence wins (same as Go).

### Column Filtering (`--cols`)

Applied before diffing. Both DataFrames are narrowed to the selected columns. Key columns are always retained, even when omitted from `--cols`.

### Tolerance

- Applies to Float64 columns and Int64-vs-Float64 cross-type comparisons in key-based mode.
- Positional mode hashes exact string values; tolerance does not apply.
- If `--tolerance` is set but no numeric columns exist, it is a silent no-op.

## Output Formats

### No Differences

When the diff result is empty (no added, removed, or modified rows):

- **Text:** prints "No differences found." to stdout.
- **Markdown:** prints "No differences found."
- **JSON:** `{"added":[],"removed":[],"modified":[]}`
- **CSV:** header row only (no data rows).

All formats exit 0.

### Text (default, colored)

```
--- Sheet1 (old.xlsx)
+++ Sheet1 (new.xlsx)

Added: 1 | Removed: 1 | Modified: 1

- ID: "3"  Name: "Charlie"  Score: "70"
+ ID: "4"  Name: "Dana"  Score: "85"

~ ID: "1"
    Score: "90" → "95"
```

ANSI colors: red (`-`), green (`+`), yellow (`~`). Auto-detects TTY via `std::io::IsTerminal`; `--no-color` overrides.

### Markdown (`--format markdown`)

```markdown
## Added (1)
| ID | Name | Score |
|----|------|-------|
| 4  | Dana | 85    |

## Removed (1)
| ID | Name | Score |
|----|------|-------|
| 3  | Charlie | 70 |

## Modified (1)
| Key (ID) | Column | Old | New |
|----------|--------|-----|-----|
| 1        | Score  | 90  | 95  |
```

Clean, uncolored tables. The added/removed tables can reuse `formatter.rs::render_table_header()` and `render_table_rows()` for consistent alignment with xlcat/xlfilter output. The modified table is diff-specific.

### JSON (`--format json`)

```json
{
  "added": [{"ID": "4", "Name": "Dana", "Score": "85"}],
  "removed": [{"ID": "3", "Name": "Charlie", "Score": "70"}],
  "modified": [
    {
      "key": {"ID": "1"},
      "changes": [{"column": "Score", "old": "90", "new": "95"}]
    }
  ]
}
```

Empty arrays are `[]`, never null. Pretty-printed with 2-space indentation.

### CSV (`--format csv`)

```csv
_status,ID,Name,Score,_old_ID,_old_Name,_old_Score
added,4,Dana,85,,,
removed,3,Charlie,70,,,
modified,1,,95,,,90
```

`_status` column indicates row type. `_old_*` columns hold previous values. For modified rows: new values go in the main columns, old values in `_old_*` columns, only for changed cells. Unchanged cells are empty in both positions. Key column values appear in the main columns to identify the row.

## Module Layout

### New Files

- **`src/diff.rs`** — Core diff logic: `DiffResult`, positional/key-based algorithms, tolerance, duplicate key detection.
- **`src/bin/xldiff.rs`** — CLI binary: clap argument parsing, sheet resolution, output formatting (text, markdown, json, csv).

### Reused Modules

- **`reader.rs`** — `read_sheet()` / `read_sheet_with_skip()` to load DataFrames.
- **`metadata.rs`** — `read_file_info()` for sheet resolution.
- **`cell.rs`** — Column letter resolution (shared with filter.rs via `col_letter_to_index()`).
- **`formatter.rs`** — `render_table_header()` and `render_table_rows()` for markdown added/removed tables.

### Output Formatting

The four format renderers live in the binary (`xldiff.rs`), not in `formatter.rs`. Diff output structure differs enough from xlcat's table rendering to warrant separation. The markdown formatter reuses `formatter.rs` table-rendering helpers for the added/removed sections.

### Cargo.toml Changes

- Add `[[bin]] name = "xldiff"` entry.
- Add `serde = { version = "1", features = ["derive"] }` and `serde_json = "1"` for JSON output.
- TTY detection via `std::io::IsTerminal` (stable since Rust 1.70, no extra crate needed).

### Pipeline

```
parse args → resolve sheets → read DataFrames
  → narrow to --cols → diff (positional or key-based)
  → format output → print → exit code
```

## Testing

- **Unit tests** in `diff.rs`: positional mode, key-based mode, tolerance (including cross-type Int64/Float64, NaN handling), duplicate keys, column filtering, edge cases (empty sheets, single-row, mismatched columns/headers).
- **Integration tests** in `tests/test_xldiff.rs`: CLI end-to-end with `assert_cmd`.
- **Fixtures**: add helpers to `tests/common/mod.rs`:
  - `create_diff_pair()` — two files with added/removed/identical rows for positional mode.
  - `create_diff_pair_with_keys()` — two files with key columns and modified cells.
  - `create_diff_pair_with_floats()` — two files with float values differing by small amounts for tolerance testing.
