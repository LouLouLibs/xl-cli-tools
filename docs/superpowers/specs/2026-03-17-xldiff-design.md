# xldiff — Excel Sheet Comparison Tool

Port of [go-xldiff](../../../go-xldiff) to Rust, leveraging polars for type-aware comparison.

## CLI Interface

```
xldiff <FILE1>[:<SHEET>] <FILE2>[:<SHEET>] [flags]
```

Two positional file arguments, each optionally suffixed with `:SheetName` or `:0` (0-based index). The same file may appear twice with different sheets (e.g., `report.xlsx:Q1 report.xlsx:Q2`).

### Flags

| Flag | Type | Default | Description |
|------|------|---------|-------------|
| `--key <COLS>` | string | *(none)* | Comma-separated key columns (name or letter). Omit for positional mode. |
| `--cols <COLS>` | string | *(all)* | Columns to compare (name or letter). Key columns always included. |
| `--skip <N>[,<M>]` | string | `0` | Rows to skip before header. Single value applies to both files; comma pair sets each independently. |
| `--no-header` | bool | false | Treat first row as data; generate synthetic `Col0`, `Col1`, … headers. |
| `--tolerance <F>` | f64 | *(none)* | Numeric tolerance for float comparisons (e.g., `0.01`). |
| `--format <FMT>` | string | `text` | Output format: `text`, `markdown`, `json`, `csv`. |
| `--no-color` | bool | false | Force-disable ANSI colors (auto-detected by default). |

### Exit Codes

- **0** — no differences
- **1** — differences found
- **2** — invalid arguments or runtime error

## Data Structures

```rust
pub struct DiffResult {
    pub headers: Vec<String>,
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
    pub values: Vec<String>,
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

## Diff Algorithm

### Positional Mode (no `--key`)

1. Hash each row: join all compared columns with a null-byte separator.
2. Build frequency maps (`HashMap<RowHash, usize>`) for both tables.
3. Walk table A: if B has remaining copies, consume one; otherwise mark as removed.
4. Walk table B: if A has remaining copies, consume one; otherwise mark as added.
5. No "modified" rows — any cell change produces a different hash.

### Key-Based Mode (`--key` provided)

1. Resolve key column names/letters to DataFrame column indices.
2. Use polars `anti_join` on key columns: A anti-join B → removed, B anti-join A → added.
3. Use polars `inner_join` on key columns → matched row pairs.
4. Compare non-key columns cell by cell for each matched pair:
   - **Numeric (Float64):** if `--tolerance` is set, equal when `|a - b| <= tolerance`; otherwise exact.
   - **Integer (Int64):** exact comparison.
   - **String:** exact string comparison.
   - **Datetime:** compare as timestamps.
   - **Boolean:** exact match.
   - **Null vs value:** treated as different.
5. Rows with any cell difference become a `ModifiedRow` with a list of `CellChange`s.
6. Duplicate keys produce a warning on stderr.

### Column Filtering (`--cols`)

Applied before diffing. Both DataFrames are narrowed to the selected columns. Key columns are always retained, even when omitted from `--cols`.

### Tolerance

Applies only to Float64 columns in key-based mode. Positional mode hashes exact values; tolerance does not apply.

## Output Formats

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

ANSI colors: red (`-`), green (`+`), yellow (`~`). Auto-detects TTY; `--no-color` overrides. When no differences exist, prints "No differences found." and exits 0.

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

Clean, uncolored tables consistent with xlcat/xlfilter output. Suitable for LLM pipelines.

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

`_status` column indicates row type. `_old_*` columns hold previous values. Modified rows populate only changed cells; unchanged cells remain empty.

## Module Layout

### New Files

- **`src/diff.rs`** — Core diff logic: `DiffResult`, positional/key-based algorithms, tolerance, duplicate key detection.
- **`src/bin/xldiff.rs`** — CLI binary: clap argument parsing, sheet resolution, output formatting (text, markdown, json, csv).

### Reused Modules

- **`reader.rs`** — `read_sheet()` / `read_sheet_with_skip()` to load DataFrames.
- **`metadata.rs`** — `read_file_info()` for sheet resolution.
- **`cell.rs`** — Column letter resolution (shared with filter.rs).

### Output Formatting

The four format renderers live in the binary (`xldiff.rs`), not in `formatter.rs`. Diff output structure differs from xlcat's table rendering; keeping them separate preserves focus.

### Cargo.toml Changes

- Add `[[bin]] name = "xldiff"` entry.
- Add `serde_json = "1"` for JSON output.
- Use `std::io::IsTerminal` (stable since Rust 1.70) for TTY detection.

### Pipeline

```
parse args → resolve sheets → read DataFrames
  → narrow to --cols → diff (positional or key-based)
  → format output → print → exit code
```

## Testing

- **Unit tests** in `diff.rs`: positional mode, key-based mode, tolerance, duplicate keys, column filtering, edge cases (empty sheets, single-row, mismatched columns).
- **Integration tests** in `tests/test_xldiff.rs`: CLI end-to-end with `assert_cmd`.
- **Fixtures**: add `create_diff_pair()` and `create_diff_pair_with_keys()` helpers to `tests/common/mod.rs`.
