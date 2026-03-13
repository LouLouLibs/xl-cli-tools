mod common;

use tempfile::TempDir;

#[test]
fn test_simple_file_metadata_header() {
    let dir = TempDir::new().unwrap();
    let path = dir.path().join("simple.xlsx");
    common::create_simple(&path);
    assert!(path.exists());
    assert!(std::fs::metadata(&path).unwrap().len() > 0);
}
