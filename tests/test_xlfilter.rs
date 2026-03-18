mod common;

use assert_cmd::Command;
use predicates::prelude::*;
use tempfile::TempDir;

fn xlfilter() -> Command {
    Command::cargo_bin("xlfilter").unwrap()
}

fn setup() -> (TempDir, std::path::PathBuf) {
    let dir = TempDir::new().unwrap();
    let path = dir.path().join("test.xlsx");
    common::create_filterable(&path);
    (dir, path)
}

// === Basic functionality ===

#[test]
fn no_flags_shows_all_rows() {
    let (_dir, path) = setup();
    xlfilter()
        .arg(&path)
        .assert()
        .success()
        .stdout(predicate::str::contains("CA"))
        .stdout(predicate::str::contains("NY"))
        .stdout(predicate::str::contains("TX"))
        .stdout(predicate::str::contains("FL"))
        .stderr(predicate::str::contains("6 rows"));
}

#[test]
fn where_eq_string() {
    let (_dir, path) = setup();
    xlfilter()
        .arg(&path)
        .args(["--where", "State=CA"])
        .assert()
        .success()
        .stdout(predicate::str::contains("Los Angeles"))
        .stdout(predicate::str::contains("San Francisco"))
        .stdout(predicate::str::contains("NY").not())
        .stderr(predicate::str::contains("2 rows"));
}

#[test]
fn where_gt_numeric() {
    let (_dir, path) = setup();
    xlfilter()
        .arg(&path)
        .args(["--where", "Amount>1500"])
        .assert()
        .success()
        .stdout(predicate::str::contains("New York"))
        .stdout(predicate::str::contains("Miami"))
        .stderr(predicate::str::contains("2 rows"));
}

#[test]
fn where_multiple_and() {
    let (_dir, path) = setup();
    xlfilter()
        .arg(&path)
        .args(["--where", "State=CA", "--where", "Amount>1000"])
        .assert()
        .success()
        .stdout(predicate::str::contains("Los Angeles"))
        .stdout(predicate::str::contains("San Francisco").not())
        .stderr(predicate::str::contains("1 rows"));
}

#[test]
fn where_not_eq() {
    let (_dir, path) = setup();
    xlfilter()
        .arg(&path)
        .args(["--where", "Status!=Draft"])
        .assert()
        .success()
        .stderr(predicate::str::contains("4 rows"));
}

#[test]
fn where_contains() {
    let (_dir, path) = setup();
    xlfilter()
        .arg(&path)
        .args(["--where", "City~angel"])
        .assert()
        .success()
        .stdout(predicate::str::contains("Los Angeles"))
        .stderr(predicate::str::contains("1 rows"));
}

#[test]
fn where_not_contains() {
    let (_dir, path) = setup();
    xlfilter()
        .arg(&path)
        .args(["--where", "Status!~raft"])
        .assert()
        .success()
        .stderr(predicate::str::contains("4 rows"));
}

#[test]
fn where_no_matches() {
    let (_dir, path) = setup();
    xlfilter()
        .arg(&path)
        .args(["--where", "State=ZZ"])
        .assert()
        .success()
        .stderr(predicate::str::contains("0 rows"));
}

// === Column selection ===

#[test]
fn cols_by_name() {
    let (_dir, path) = setup();
    xlfilter()
        .arg(&path)
        .args(["--cols", "State,Amount"])
        .assert()
        .success()
        .stdout(predicate::str::contains("State"))
        .stdout(predicate::str::contains("Amount"))
        .stdout(predicate::str::contains("City").not())
        .stdout(predicate::str::contains("Year").not());
}

#[test]
fn cols_by_letter() {
    let (_dir, path) = setup();
    xlfilter()
        .arg(&path)
        .args(["--cols", "A,C"])
        .assert()
        .success()
        .stdout(predicate::str::contains("State"))
        .stdout(predicate::str::contains("Amount"))
        .stdout(predicate::str::contains("City").not());
}

#[test]
fn cols_mixed_letter_and_name() {
    let (_dir, path) = setup();
    xlfilter()
        .arg(&path)
        .args(["--cols", "A,Amount"])
        .assert()
        .success()
        .stdout(predicate::str::contains("State"))
        .stdout(predicate::str::contains("Amount"))
        .stdout(predicate::str::contains("City").not());
}

// === Sort ===

#[test]
fn sort_desc() {
    let (_dir, path) = setup();
    xlfilter()
        .arg(&path)
        .args(["--sort", "Amount:desc", "--cols", "City,Amount"])
        .assert()
        .success()
        .stdout(predicate::str::contains("Miami")); // 3000 = highest
}

#[test]
fn sort_asc() {
    let (_dir, path) = setup();
    xlfilter()
        .arg(&path)
        .args(["--sort", "Amount:asc", "--limit", "1", "--cols", "City,Amount"])
        .assert()
        .success()
        .stdout(predicate::str::contains("Albany")); // 500 = lowest
}

#[test]
fn sort_by_column_letter() {
    let (_dir, path) = setup();
    // C = Amount column
    xlfilter()
        .arg(&path)
        .args(["--sort", "C:desc", "--limit", "1", "--cols", "City,Amount"])
        .assert()
        .success()
        .stdout(predicate::str::contains("Miami")); // 3000 = highest
}

// === Limit, head, tail ===

#[test]
fn limit_caps_output() {
    let (_dir, path) = setup();
    xlfilter()
        .arg(&path)
        .args(["--limit", "3"])
        .assert()
        .success()
        .stderr(predicate::str::contains("3 rows"));
}

#[test]
fn head_before_filter() {
    let (_dir, path) = setup();
    // First 3 rows: CA/LA, NY/NYC, CA/SF
    // Filter State=NY → only NYC
    xlfilter()
        .arg(&path)
        .args(["--head", "3", "--where", "State=NY"])
        .assert()
        .success()
        .stderr(predicate::str::contains("1 rows"));
}

#[test]
fn tail_before_filter() {
    let (_dir, path) = setup();
    // Last 3 rows: TX/Houston, NY/Albany, FL/Miami
    // Filter State=NY → only Albany
    xlfilter()
        .arg(&path)
        .args(["--tail", "3", "--where", "State=NY"])
        .assert()
        .success()
        .stderr(predicate::str::contains("1 rows"));
}

#[test]
fn head_and_tail_mutually_exclusive() {
    let (_dir, path) = setup();
    xlfilter()
        .arg(&path)
        .args(["--head", "3", "--tail", "3"])
        .assert()
        .failure()
        .stderr(predicate::str::contains("mutually exclusive"));
}

// === CSV output ===

#[test]
fn csv_output() {
    let (_dir, path) = setup();
    xlfilter()
        .arg(&path)
        .args(["--csv", "--where", "State=CA"])
        .assert()
        .success()
        .stdout(predicate::str::contains(",")) // CSV has commas
        .stdout(predicate::str::contains("|").not()); // no markdown pipes
}

// === Sheet selection ===

#[test]
fn sheet_by_name() {
    let (_dir, path) = setup();
    xlfilter()
        .arg(&path)
        .args(["--sheet", "Data"])
        .assert()
        .success();
}

#[test]
fn sheet_not_found() {
    let (_dir, path) = setup();
    xlfilter()
        .arg(&path)
        .args(["--sheet", "NoSuchSheet"])
        .assert()
        .failure()
        .stderr(predicate::str::contains("not found"));
}

// === Error cases ===

#[test]
fn file_not_found() {
    xlfilter()
        .arg("nonexistent.xlsx")
        .assert()
        .failure()
        .stderr(predicate::str::contains("not found"));
}

#[test]
fn bad_filter_expr() {
    let (_dir, path) = setup();
    xlfilter()
        .arg(&path)
        .args(["--where", "NoOperator"])
        .assert()
        .failure()
        .stderr(predicate::str::contains("no operator found"));
}

#[test]
fn bad_sort_dir() {
    let (_dir, path) = setup();
    xlfilter()
        .arg(&path)
        .args(["--sort", "Amount:up"])
        .assert()
        .failure()
        .stderr(predicate::str::contains("invalid sort direction"));
}

#[test]
fn unknown_column_in_where() {
    let (_dir, path) = setup();
    xlfilter()
        .arg(&path)
        .args(["--where", "Foo=bar"])
        .assert()
        .failure()
        .stderr(predicate::str::contains("not found"));
}

#[test]
fn unknown_column_in_cols() {
    let (_dir, path) = setup();
    xlfilter()
        .arg(&path)
        .args(["--cols", "State,Foo"])
        .assert()
        .failure()
        .stderr(predicate::str::contains("not found"));
}

#[test]
fn gt_with_non_numeric_value() {
    let (_dir, path) = setup();
    xlfilter()
        .arg(&path)
        .args(["--where", "Amount>abc"])
        .assert()
        .failure()
        .stderr(predicate::str::contains("numeric value"));
}
