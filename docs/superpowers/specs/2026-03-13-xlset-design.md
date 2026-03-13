# xlset — Excel Cell Writer

## Purpose

A Rust CLI tool that modifies cells in existing `.xlsx` files using
`umya-spreadsheet`. Preserves all formatting, formulas, charts, and
structure it doesn't touch. Lives in the same repo as `xlcat` (the
read-only viewer) but is a separate binary.

## CLI Interface

```
xlset <file> <cell>=<value> [<cell>=<value> ...]
xlset <file> --from <csv-file|->
xlset <file> --sheet "Revenue" A1=42 B2="hello"
xlset <file> --output other.xlsx A1=42
```

### Flags

| Flag | Purpose |
|------|---------|
| `--sheet <name\|index>` | Target sheet (default: first sheet). Name or 0-based index. |
| `--output <path>` | Write to a new file instead of modifying in-place. |
| `--from <path\|->` | Read cell assignments from a CSV file or stdin (`-`). |

## Cell Assignment Syntax

Positional arguments after the file path are cell assignments:

```
A1=42              # auto-infer: integer
B2=3.14            # auto-infer: float
C3=true            # auto-infer: boolean
D4=2024-01-15      # auto-infer: date
E5=hello world     # auto-infer: string
F6:str=07401       # explicit: force string (preserves leading zero)
G7:num=42          # explicit: force number
H8:bool=1          # explicit: force boolean
I9:date=2024-01-15 # explicit: force date
```

### Type Tags

Optional type tags override auto-inference: `:str`, `:num`, `:bool`, `:date`.
Placed between the cell reference and the `=` sign.

### Auto-Inference Rules (in order)

1. If explicit type tag is present → use that type.
2. `true` or `false` (case-insensitive) → boolean.
3. Parseable as `i64` (no decimal point) → integer (stored as f64 in Excel).
4. Parseable as `f64` → float.
5. Matches `YYYY-MM-DD` pattern → date (stored as Excel serial number
   with `yyyy-mm-dd` format code applied to the cell).
6. Empty value after `=` (e.g., `A1=`) → empty string.
7. Everything else → string.

## `--from` CSV Format

```csv
cell,value
A1,42
B2,hello
C3:str,07401
```

- First row is skipped if the first field does not parse as a valid cell
  reference (e.g., a header like `cell,value` or `ref,val`).
- Each row is a `cell,value` pair. Standard RFC 4180 CSV quoting applies —
  values containing commas must be quoted: `A1,"hello, world"`.
- Type tags work in the cell column: `C3:str,07401`.
- Reads from a file path or `-` for stdin.
- Can be combined with positional args. **Duplicate resolution:** last-write-wins,
  with positional args applied after CSV. If the CSV sets `A1=42` and a
  positional arg sets `A1=99`, the cell gets `99`.

## Cell Address Parsing

Supports standard Excel A1 notation:

- Column: `A` through `XFD` (1 to 16384).
- Row: `1` through `1048576`.
- Examples: `A1`, `Z99`, `AA1`, `XFD1048576`.
- Case-insensitive: `a1` = `A1`.

Passed to umya-spreadsheet using its string-based cell reference API
(e.g., `worksheet.get_cell_mut("A1")`).

## Sheet Selection

- `--sheet "Revenue"` selects by name.
- `--sheet 0` selects by 0-based index.
- If the argument matches a sheet name exactly, the name takes precedence
  over interpreting it as an index (same behavior as xlcat).
- Default: first sheet in the workbook.
- If sheet not found: error listing available sheets.

## Output Behavior

- **Default:** modifies file in-place.
- **`--output path.xlsx`:** writes to a new file, original untouched.
- Errors are reported before any write occurs when possible.

## Stderr Confirmation

On success, prints to stderr:
```
xlset: updated 3 cells in Revenue (report.xlsx)
```

No output to stdout. This keeps xlset pipe-friendly.

## Exit Codes

| Code | Meaning |
|------|---------|
| 0 | Success |
| 1 | Runtime error (file not found, corrupt, sheet missing, write failure) |
| 2 | Invalid arguments (bad cell reference, bad type tag, bad CSV) |

All error messages go to stderr.

## Technology

- **Language:** Rust
- **Excel read/write:** `umya-spreadsheet` (round-trip editing, preserves formatting)
- **CLI parsing:** `clap` (derive API, shared with xlcat)
- **Cell parsing / type inference:** shared library code with xlcat
- **No runtime dependencies** — single compiled binary

## Project Structure

Same repo, two binaries sharing a library:

```
Cargo.toml            # [[bin]] entries for xlcat and xlset
src/
├── lib.rs            # crate library root, re-exports modules
├── bin/
│   ├── xlcat.rs      # xlcat entry point (moved from main.rs)
│   └── xlset.rs      # xlset entry point
├── cell.rs           # A1 notation parser, value type inference
├── writer.rs         # umya-spreadsheet write logic
├── metadata.rs       # existing (used by xlcat)
├── reader.rs         # existing (used by xlcat)
└── formatter.rs      # existing (used by xlcat)
```

`umya-spreadsheet` is added as a dependency. Existing calamine/polars deps
remain for xlcat. Shared code in `lib.rs` and `cell.rs` must not import
crate-specific types (polars, umya-spreadsheet) unconditionally — only
the binary entry points and their dedicated modules should pull those in.
With LTO enabled, the linker strips unused code per binary.

## Error Handling

- **File does not exist:** error. xlset operates on existing files only; it
  does not create new workbooks.
- **File is `.xls` (not `.xlsx`):** "xlset only supports .xlsx files."
- **File not found:** clear error with path.
- **File not writable:** "Cannot write to: <path>" with OS error.
- **Invalid cell reference:** "Invalid cell reference: <ref>" with hint
  (e.g., "Expected format: A1, B2, AA100").
- **Invalid type tag:** "Unknown type tag: <tag>. Valid tags: str, num, bool, date".
- **Sheet not found:** list available sheets in error message.
- **CSV parse error:** "Error on line N: <detail>".
- **No assignments:** "No cell assignments provided. Use positional args or --from."

## Future Possibilities (Not in Scope)

- Formula setting (`A1==SUM(B1:B10)`)
- Range fill (`A1:A10=0`)
- Cell deletion / clearing
- Row/column insertion
- Conditional updates (`--if-empty`)
