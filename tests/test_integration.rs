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

#[test]
fn test_large_file_gate_triggers() {
    let dir = TempDir::new().unwrap();
    let path = dir.path().join("many.xlsx");
    common::create_many_rows(&path);

    // Set max-size very low so gate triggers
    xlcat()
        .arg(path.to_str().unwrap())
        .arg("--max-size")
        .arg("1K")
        .assert()
        .success()
        .stdout(predicate::str::contains("Large file"));
}

#[test]
fn test_large_file_gate_overridden_by_head() {
    let dir = TempDir::new().unwrap();
    let path = dir.path().join("many.xlsx");
    common::create_many_rows(&path);

    let output = xlcat()
        .arg(path.to_str().unwrap())
        .arg("--max-size")
        .arg("1K")
        .arg("--head")
        .arg("5")
        .assert()
        .success();

    let stdout = String::from_utf8(output.get_output().stdout.clone()).unwrap();
    assert!(!stdout.contains("Large file"));
}

#[test]
fn test_large_file_gate_overridden_by_all() {
    let dir = TempDir::new().unwrap();
    let path = dir.path().join("many.xlsx");
    common::create_many_rows(&path);

    let output = xlcat()
        .arg(path.to_str().unwrap())
        .arg("--max-size")
        .arg("1K")
        .arg("--all")
        .assert()
        .success();

    let stdout = String::from_utf8(output.get_output().stdout.clone()).unwrap();
    assert!(!stdout.contains("Large file"));
}

#[test]
fn test_empty_data_headers_only() {
    let dir = TempDir::new().unwrap();
    let path = dir.path().join("empty_data.xlsx");
    common::create_empty_data(&path);

    xlcat()
        .arg(path.to_str().unwrap())
        .assert()
        .success()
        .stdout(predicate::str::contains("no data rows"));
}

#[test]
fn test_head_and_tail_together() {
    let dir = TempDir::new().unwrap();
    let path = dir.path().join("many.xlsx");
    common::create_many_rows(&path);

    xlcat()
        .arg(path.to_str().unwrap())
        .arg("--head")
        .arg("3")
        .arg("--tail")
        .arg("2")
        .assert()
        .success()
        .stdout(predicate::str::contains("omitted"));
}

#[test]
fn test_tail_flag_alone() {
    let dir = TempDir::new().unwrap();
    let path = dir.path().join("many.xlsx");
    common::create_many_rows(&path);

    let output = xlcat()
        .arg(path.to_str().unwrap())
        .arg("--tail")
        .arg("3")
        .assert()
        .success();

    let stdout = String::from_utf8(output.get_output().stdout.clone()).unwrap();
    assert!(!stdout.contains("omitted"));
    // Should contain the last rows (ids near 80)
    assert!(stdout.contains("80"));
}

#[test]
fn test_csv_respects_head() {
    let dir = TempDir::new().unwrap();
    let path = dir.path().join("many.xlsx");
    common::create_many_rows(&path);

    let output = xlcat()
        .arg(path.to_str().unwrap())
        .arg("--csv")
        .arg("--head")
        .arg("3")
        .assert()
        .success();

    let stdout = String::from_utf8(output.get_output().stdout.clone()).unwrap();
    // Header + 3 data rows = 4 non-empty lines
    let lines: Vec<&str> = stdout.trim().lines().collect();
    assert_eq!(lines.len(), 4, "Expected header + 3 data rows, got: {}", stdout);
}

#[test]
fn test_head_tail_overlap_shows_all() {
    let dir = TempDir::new().unwrap();
    let path = dir.path().join("simple.xlsx");
    common::create_simple(&path);

    // 5 rows, head 3 + tail 3 = 6 > 5, so show all without duplication
    let output = xlcat()
        .arg(path.to_str().unwrap())
        .arg("--head")
        .arg("3")
        .arg("--tail")
        .arg("3")
        .assert()
        .success();

    let stdout = String::from_utf8(output.get_output().stdout.clone()).unwrap();
    assert!(!stdout.contains("omitted"));
    assert!(stdout.contains("Alice"));
    assert!(stdout.contains("Eve"));
}

#[test]
fn test_wrong_extension() {
    let dir = TempDir::new().unwrap();
    let path = dir.path().join("data.csv");
    std::fs::write(&path, "a,b\n1,2\n").unwrap();

    xlcat()
        .arg(path.to_str().unwrap())
        .assert()
        .failure()
        .stderr(predicate::str::contains("Expected .xls or .xlsx"));
}

#[test]
fn test_sheet_not_found() {
    let dir = TempDir::new().unwrap();
    let path = dir.path().join("simple.xlsx");
    common::create_simple(&path);

    xlcat()
        .arg(path.to_str().unwrap())
        .arg("--sheet")
        .arg("Nonexistent")
        .assert()
        .failure()
        .stderr(predicate::str::contains("not found"));
}

#[test]
fn test_exit_code_success() {
    let dir = TempDir::new().unwrap();
    let path = dir.path().join("simple.xlsx");
    common::create_simple(&path);

    xlcat()
        .arg(path.to_str().unwrap())
        .assert()
        .code(0);
}

#[test]
fn test_exit_code_runtime_error() {
    xlcat()
        .arg("/nonexistent.xlsx")
        .assert()
        .code(1);
}

#[test]
fn test_exit_code_invalid_args() {
    let dir = TempDir::new().unwrap();
    let path = dir.path().join("simple.xlsx");
    common::create_simple(&path);

    xlcat()
        .arg(path.to_str().unwrap())
        .arg("--schema")
        .arg("--describe")
        .assert()
        .code(2);
}

#[test]
fn test_describe_multi_sheet_no_sheet_flag() {
    let dir = TempDir::new().unwrap();
    let path = dir.path().join("multi.xlsx");
    common::create_multi_sheet(&path);

    xlcat()
        .arg(path.to_str().unwrap())
        .arg("--describe")
        .assert()
        .success()
        .stdout(predicate::str::contains("## Sheet: Revenue"))
        .stdout(predicate::str::contains("## Sheet: Expenses"))
        .stdout(predicate::str::contains("count"))
        .stdout(predicate::str::contains("mean"));
}

#[test]
fn test_csv_multi_sheet_without_sheet_is_error() {
    let dir = TempDir::new().unwrap();
    let path = dir.path().join("multi.xlsx");
    common::create_multi_sheet(&path);

    xlcat()
        .arg(path.to_str().unwrap())
        .arg("--csv")
        .assert()
        .failure()
        .stderr(predicate::str::contains("Multiple sheets"));
}
