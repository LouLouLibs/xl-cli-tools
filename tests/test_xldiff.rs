mod common;

use assert_cmd::Command;
use predicates::prelude::*;
use tempfile::TempDir;

fn xldiff() -> Command {
    Command::cargo_bin("xldiff").unwrap()
}

// ---------------------------------------------------------------------------
// test_no_diff
// ---------------------------------------------------------------------------

#[test]
fn test_no_diff() {
    let dir = TempDir::new().unwrap();
    let path_a = dir.path().join("a.xlsx");
    let path_b = dir.path().join("b.xlsx");
    common::create_simple(&path_a);
    common::create_simple(&path_b);

    xldiff()
        .arg(&path_a)
        .arg(&path_b)
        .arg("--no-color")
        .assert()
        .success()
        .stdout(predicate::str::contains("No differences found."));
}

// ---------------------------------------------------------------------------
// test_positional_diff
// ---------------------------------------------------------------------------

#[test]
fn test_positional_diff() {
    let dir = TempDir::new().unwrap();
    let path_a = dir.path().join("a.xlsx");
    let path_b = dir.path().join("b.xlsx");
    common::create_diff_pair(&path_a, &path_b);

    xldiff()
        .arg(&path_a)
        .arg(&path_b)
        .arg("--no-color")
        .assert()
        .code(1)
        .stdout(predicate::str::contains("Removed: 1"))
        .stdout(predicate::str::contains("Added: 1"))
        .stdout(predicate::str::contains("Bob"));
}

// ---------------------------------------------------------------------------
// test_keyed_diff
// ---------------------------------------------------------------------------

#[test]
fn test_keyed_diff() {
    let dir = TempDir::new().unwrap();
    let path_a = dir.path().join("a.xlsx");
    let path_b = dir.path().join("b.xlsx");
    common::create_diff_pair_with_keys(&path_a, &path_b);

    xldiff()
        .arg(&path_a)
        .arg(&path_b)
        .args(["--key", "ID", "--no-color"])
        .assert()
        .code(1)
        .stdout(predicate::str::contains("Modified: 1"))
        .stdout(predicate::str::contains("Removed: 1"))
        .stdout(predicate::str::contains("Added: 1"));
}

// ---------------------------------------------------------------------------
// test_tolerance
// ---------------------------------------------------------------------------

#[test]
fn test_tolerance() {
    let dir = TempDir::new().unwrap();
    let path_a = dir.path().join("a.xlsx");
    let path_b = dir.path().join("b.xlsx");
    common::create_diff_pair_with_floats(&path_a, &path_b);

    // With tolerance 0.01: ID=1 diff is 0.001 (within tolerance), ID=2 diff is 0.1 (exceeds)
    xldiff()
        .arg(&path_a)
        .arg(&path_b)
        .args(["--key", "ID", "--tolerance", "0.01", "--no-color"])
        .assert()
        .code(1)
        .stdout(predicate::str::contains("Modified: 1"));
}

// ---------------------------------------------------------------------------
// test_json_format
// ---------------------------------------------------------------------------

#[test]
fn test_json_format() {
    let dir = TempDir::new().unwrap();
    let path_a = dir.path().join("a.xlsx");
    let path_b = dir.path().join("b.xlsx");
    common::create_diff_pair_with_keys(&path_a, &path_b);

    xldiff()
        .arg(&path_a)
        .arg(&path_b)
        .args(["--key", "ID", "--format", "json"])
        .assert()
        .code(1)
        .stdout(predicate::str::contains("\"added\""))
        .stdout(predicate::str::contains("\"removed\""))
        .stdout(predicate::str::contains("\"modified\""));
}

// ---------------------------------------------------------------------------
// test_markdown_format
// ---------------------------------------------------------------------------

#[test]
fn test_markdown_format() {
    let dir = TempDir::new().unwrap();
    let path_a = dir.path().join("a.xlsx");
    let path_b = dir.path().join("b.xlsx");
    common::create_diff_pair_with_keys(&path_a, &path_b);

    xldiff()
        .arg(&path_a)
        .arg(&path_b)
        .args(["--key", "ID", "--format", "markdown"])
        .assert()
        .code(1)
        .stdout(predicate::str::contains("## Added"))
        .stdout(predicate::str::contains("## Removed"))
        .stdout(predicate::str::contains("## Modified"));
}

// ---------------------------------------------------------------------------
// test_csv_format
// ---------------------------------------------------------------------------

#[test]
fn test_csv_format() {
    let dir = TempDir::new().unwrap();
    let path_a = dir.path().join("a.xlsx");
    let path_b = dir.path().join("b.xlsx");
    common::create_diff_pair_with_keys(&path_a, &path_b);

    xldiff()
        .arg(&path_a)
        .arg(&path_b)
        .args(["--key", "ID", "--format", "csv"])
        .assert()
        .code(1)
        .stdout(predicate::str::contains("_status"))
        .stdout(predicate::str::contains("added"))
        .stdout(predicate::str::contains("removed"))
        .stdout(predicate::str::contains("modified"));
}

// ---------------------------------------------------------------------------
// test_file_not_found
// ---------------------------------------------------------------------------

#[test]
fn test_file_not_found() {
    xldiff()
        .arg("nonexistent_a.xlsx")
        .arg("nonexistent_b.xlsx")
        .assert()
        .code(2)
        .stderr(predicate::str::contains("file not found"));
}

// ---------------------------------------------------------------------------
// test_sheet_selector
// ---------------------------------------------------------------------------

#[test]
fn test_sheet_selector() {
    let dir = TempDir::new().unwrap();
    let path = dir.path().join("multi.xlsx");
    common::create_multi_sheet(&path);

    // Revenue and Expenses sheets have different schemas and data
    let arg_revenue = format!("{}:Revenue", path.display());
    let arg_expenses = format!("{}:Expenses", path.display());

    xldiff()
        .arg(&arg_revenue)
        .arg(&arg_expenses)
        .assert()
        .code(1);
}

// ---------------------------------------------------------------------------
// test_cols_filter
// ---------------------------------------------------------------------------

#[test]
fn test_cols_filter() {
    let dir = TempDir::new().unwrap();
    let path_a = dir.path().join("a.xlsx");
    let path_b = dir.path().join("b.xlsx");
    common::create_diff_pair_with_keys(&path_a, &path_b);

    // With --cols ID,Name: Score column is excluded from comparison.
    // ID=1: Score changed 90→95, but Score is excluded, so Name is same → no modification.
    // ID=3: removed, ID=4: added.
    // Modified should be 0.
    xldiff()
        .arg(&path_a)
        .arg(&path_b)
        .args(["--key", "ID", "--cols", "ID,Name", "--no-color"])
        .assert()
        .code(1)
        .stdout(predicate::str::contains("Modified: 0"))
        .stdout(predicate::str::contains("Added: 1"))
        .stdout(predicate::str::contains("Removed: 1"));
}
