mod common;

use assert_cmd::Command;
use predicates::prelude::*;
use tempfile::TempDir;

fn xlcat() -> Command {
    Command::cargo_bin("xlcat").unwrap()
}

#[test]
fn test_simple_file_default() {
    let dir = TempDir::new().unwrap();
    let path = dir.path().join("simple.xlsx");
    common::create_simple(&path);

    xlcat()
        .arg(path.to_str().unwrap())
        .assert()
        .success()
        .stdout(predicate::str::contains("# File:"))
        .stdout(predicate::str::contains("# Sheets: 1"))
        .stdout(predicate::str::contains("## Sheet: Data"))
        .stdout(predicate::str::contains("| name |"))
        .stdout(predicate::str::contains("| Alice |"));
}

#[test]
fn test_schema_mode() {
    let dir = TempDir::new().unwrap();
    let path = dir.path().join("simple.xlsx");
    common::create_simple(&path);

    xlcat()
        .arg("--schema")
        .arg(path.to_str().unwrap())
        .assert()
        .success()
        .stdout(predicate::str::contains("| Column | Type |"))
        .stdout(predicate::str::contains("| name |"))
        .stdout(predicate::str::contains("| Alice |").not());
}

#[test]
fn test_multi_sheet_default_lists_schemas() {
    let dir = TempDir::new().unwrap();
    let path = dir.path().join("multi.xlsx");
    common::create_multi_sheet(&path);

    xlcat()
        .arg(path.to_str().unwrap())
        .assert()
        .success()
        .stdout(predicate::str::contains("# Sheets: 3"))
        .stdout(predicate::str::contains("## Sheet: Revenue"))
        .stdout(predicate::str::contains("## Sheet: Expenses"))
        .stdout(predicate::str::contains("## Sheet: Summary"))
        .stdout(predicate::str::contains("Use --sheet"));
}

#[test]
fn test_multi_sheet_select_by_name() {
    let dir = TempDir::new().unwrap();
    let path = dir.path().join("multi.xlsx");
    common::create_multi_sheet(&path);

    xlcat()
        .arg("--sheet")
        .arg("Revenue")
        .arg(path.to_str().unwrap())
        .assert()
        .success()
        .stdout(predicate::str::contains("| region |"))
        .stdout(predicate::str::contains("| Region 1 |"));
}

#[test]
fn test_multi_sheet_select_by_index() {
    let dir = TempDir::new().unwrap();
    let path = dir.path().join("multi.xlsx");
    common::create_multi_sheet(&path);

    xlcat()
        .arg("--sheet")
        .arg("1")
        .arg(path.to_str().unwrap())
        .assert()
        .success()
        .stdout(predicate::str::contains("## Sheet: Expenses"));
}

#[test]
fn test_head_tail_adaptive() {
    let dir = TempDir::new().unwrap();
    let path = dir.path().join("many.xlsx");
    common::create_many_rows(&path);

    xlcat()
        .arg(path.to_str().unwrap())
        .assert()
        .success()
        .stdout(predicate::str::contains("rows omitted"));
}

#[test]
fn test_head_flag() {
    let dir = TempDir::new().unwrap();
    let path = dir.path().join("many.xlsx");
    common::create_many_rows(&path);

    xlcat()
        .arg("--head")
        .arg("3")
        .arg(path.to_str().unwrap())
        .assert()
        .success()
        .stdout(predicate::str::contains("omitted").not());
}

#[test]
fn test_csv_mode() {
    let dir = TempDir::new().unwrap();
    let path = dir.path().join("simple.xlsx");
    common::create_simple(&path);

    xlcat()
        .arg("--csv")
        .arg(path.to_str().unwrap())
        .assert()
        .success()
        .stdout(predicate::str::contains("# File:").not())
        .stdout(predicate::str::contains("name,"));
}

#[test]
fn test_invalid_flag_combo() {
    let dir = TempDir::new().unwrap();
    let path = dir.path().join("simple.xlsx");
    common::create_simple(&path);

    xlcat()
        .arg("--schema")
        .arg("--head")
        .arg("10")
        .arg(path.to_str().unwrap())
        .assert()
        .code(2)
        .stderr(predicate::str::contains("cannot be combined"));
}

#[test]
fn test_file_not_found() {
    xlcat()
        .arg("/nonexistent.xlsx")
        .assert()
        .failure()
        .stderr(predicate::str::contains("Cannot"));
}

#[test]
fn test_empty_sheet() {
    let dir = TempDir::new().unwrap();
    let path = dir.path().join("empty.xlsx");
    common::create_empty_sheet(&path);

    xlcat()
        .arg(path.to_str().unwrap())
        .assert()
        .success()
        .stdout(predicate::str::contains("empty"));
}

#[test]
fn test_describe_mode() {
    let dir = TempDir::new().unwrap();
    let path = dir.path().join("simple.xlsx");
    common::create_simple(&path);

    xlcat()
        .arg(path.to_str().unwrap())
        .arg("--describe")
        .assert()
        .success()
        .stdout(predicate::str::contains("count"))
        .stdout(predicate::str::contains("mean"));
}

#[test]
fn test_all_without_sheet_on_multi() {
    let dir = TempDir::new().unwrap();
    let path = dir.path().join("multi.xlsx");
    common::create_multi_sheet(&path);

    xlcat()
        .arg("--all")
        .arg(path.to_str().unwrap())
        .assert()
        .failure()
        .stderr(predicate::str::contains("Multiple sheets"));
}
