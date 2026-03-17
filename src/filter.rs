#[derive(Debug, Clone, PartialEq)]
pub enum FilterOp {
    Eq,
    NotEq,
    Gt,
    Lt,
    Gte,
    Lte,
    Contains,
    NotContains,
}

#[derive(Debug, Clone)]
pub struct FilterExpr {
    pub column: String,
    pub op: FilterOp,
    pub value: String,
}

#[derive(Debug, Clone)]
pub struct SortSpec {
    pub column: String,
    pub descending: bool,
}

/// Parse a filter expression like "State=CA", "Amount>1000", "Name~john".
/// Scans left-to-right for the first operator character (= ! > < ~),
/// then determines the full operator.
pub fn parse_filter_expr(s: &str) -> Result<FilterExpr, String> {
    let op_chars = ['=', '!', '>', '<', '~'];
    let pos = s
        .find(|c: char| op_chars.contains(&c))
        .ok_or_else(|| {
            format!(
                "no operator found in '{}'. Use =, !=, >, <, >=, <=, ~ or !~",
                s
            )
        })?;
    if pos == 0 {
        return Err(format!("missing column name in '{}'", s));
    }
    let column = s[..pos].to_string();
    let rest = &s[pos..];
    let (op, op_len) = if rest.starts_with(">=") {
        (FilterOp::Gte, 2)
    } else if rest.starts_with("<=") {
        (FilterOp::Lte, 2)
    } else if rest.starts_with("!=") {
        (FilterOp::NotEq, 2)
    } else if rest.starts_with("!~") {
        (FilterOp::NotContains, 2)
    } else if rest.starts_with('>') {
        (FilterOp::Gt, 1)
    } else if rest.starts_with('<') {
        (FilterOp::Lt, 1)
    } else if rest.starts_with('=') {
        (FilterOp::Eq, 1)
    } else if rest.starts_with('~') {
        (FilterOp::Contains, 1)
    } else {
        return Err(format!("invalid operator in '{}'", s));
    };
    let value = rest[op_len..].to_string();
    Ok(FilterExpr { column, op, value })
}

/// Parse a sort spec like "Amount:desc" or "Name" (default asc).
/// Splits on the last colon so column names containing colons are supported.
pub fn parse_sort_spec(s: &str) -> Result<SortSpec, String> {
    if let Some(colon_pos) = s.rfind(':') {
        let col = &s[..colon_pos];
        let dir = &s[colon_pos + 1..];
        match dir.to_lowercase().as_str() {
            "asc" => Ok(SortSpec {
                column: col.to_string(),
                descending: false,
            }),
            "desc" => Ok(SortSpec {
                column: col.to_string(),
                descending: true,
            }),
            _ => Err(format!(
                "invalid sort direction '{}'. Use 'asc' or 'desc'",
                dir
            )),
        }
    } else {
        Ok(SortSpec {
            column: s.to_string(),
            descending: false,
        })
    }
}

/// Convert a column letter like "A", "B", "AA" to a 0-based index.
/// Returns None if the string isn't purely alphabetic or is empty.
fn col_letter_to_index(s: &str) -> Option<usize> {
    if s.is_empty() || !s.chars().all(|c| c.is_ascii_alphabetic()) {
        return None;
    }
    let mut idx: usize = 0;
    for c in s.to_uppercase().chars() {
        idx = idx * 26 + (c as usize - 'A' as usize + 1);
    }
    Some(idx - 1)
}

/// Resolve a column specifier to a DataFrame column name.
/// Accepts either:
/// - A header name (exact match first, then case-insensitive)
/// - A column letter like "A", "B", "AA" (mapped by position)
/// Header name match takes priority over column letter interpretation.
pub fn resolve_column(spec: &str, df_columns: &[String]) -> Result<String, String> {
    // 1. Exact header name match
    if df_columns.contains(&spec.to_string()) {
        return Ok(spec.to_string());
    }
    // 2. Case-insensitive header name match
    let spec_lower = spec.to_lowercase();
    for col in df_columns {
        if col.to_lowercase() == spec_lower {
            return Ok(col.clone());
        }
    }
    // 3. Column letter (A=0, B=1, ...) — only if purely alphabetic
    if let Some(idx) = col_letter_to_index(spec) {
        if idx < df_columns.len() {
            return Ok(df_columns[idx].clone());
        }
    }
    let available = df_columns.join(", ");
    Err(format!("column '{}' not found. Available columns: {}", spec, available))
}

/// Resolve a list of column specifiers to DataFrame column names.
pub fn resolve_columns(specs: &[String], df_columns: &[String]) -> Result<Vec<String>, String> {
    specs.iter().map(|s| resolve_column(s, df_columns)).collect()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parse_eq() {
        let expr = parse_filter_expr("State=CA").unwrap();
        assert_eq!(expr.column, "State");
        assert_eq!(expr.op, FilterOp::Eq);
        assert_eq!(expr.value, "CA");
    }

    #[test]
    fn parse_not_eq() {
        let expr = parse_filter_expr("Status!=Draft").unwrap();
        assert_eq!(expr.column, "Status");
        assert_eq!(expr.op, FilterOp::NotEq);
        assert_eq!(expr.value, "Draft");
    }

    #[test]
    fn parse_gt() {
        let expr = parse_filter_expr("Amount>1000").unwrap();
        assert_eq!(expr.column, "Amount");
        assert_eq!(expr.op, FilterOp::Gt);
        assert_eq!(expr.value, "1000");
    }

    #[test]
    fn parse_lt() {
        let expr = parse_filter_expr("Year<2024").unwrap();
        assert_eq!(expr.column, "Year");
        assert_eq!(expr.op, FilterOp::Lt);
        assert_eq!(expr.value, "2024");
    }

    #[test]
    fn parse_gte() {
        let expr = parse_filter_expr("Score>=90").unwrap();
        assert_eq!(expr.column, "Score");
        assert_eq!(expr.op, FilterOp::Gte);
        assert_eq!(expr.value, "90");
    }

    #[test]
    fn parse_lte() {
        let expr = parse_filter_expr("Price<=50.5").unwrap();
        assert_eq!(expr.column, "Price");
        assert_eq!(expr.op, FilterOp::Lte);
        assert_eq!(expr.value, "50.5");
    }

    #[test]
    fn parse_contains() {
        let expr = parse_filter_expr("Name~john").unwrap();
        assert_eq!(expr.column, "Name");
        assert_eq!(expr.op, FilterOp::Contains);
        assert_eq!(expr.value, "john");
    }

    #[test]
    fn parse_not_contains() {
        let expr = parse_filter_expr("Name!~draft").unwrap();
        assert_eq!(expr.column, "Name");
        assert_eq!(expr.op, FilterOp::NotContains);
        assert_eq!(expr.value, "draft");
    }

    #[test]
    fn parse_value_with_equals() {
        let expr = parse_filter_expr("Formula=A+B=C").unwrap();
        assert_eq!(expr.column, "Formula");
        assert_eq!(expr.op, FilterOp::Eq);
        assert_eq!(expr.value, "A+B=C");
    }

    #[test]
    fn parse_empty_value() {
        let expr = parse_filter_expr("Status=").unwrap();
        assert_eq!(expr.column, "Status");
        assert_eq!(expr.op, FilterOp::Eq);
        assert_eq!(expr.value, "");
    }

    #[test]
    fn parse_no_operator_is_err() {
        assert!(parse_filter_expr("JustAWord").is_err());
    }

    #[test]
    fn parse_no_column_is_err() {
        assert!(parse_filter_expr("=value").is_err());
    }

    #[test]
    fn parse_sort_desc() {
        let spec = parse_sort_spec("Amount:desc").unwrap();
        assert_eq!(spec.column, "Amount");
        assert!(spec.descending);
    }

    #[test]
    fn parse_sort_asc() {
        let spec = parse_sort_spec("Name:asc").unwrap();
        assert_eq!(spec.column, "Name");
        assert!(!spec.descending);
    }

    #[test]
    fn parse_sort_default_asc() {
        let spec = parse_sort_spec("Name").unwrap();
        assert_eq!(spec.column, "Name");
        assert!(!spec.descending);
    }

    #[test]
    fn parse_sort_bad_dir_is_err() {
        assert!(parse_sort_spec("Name:up").is_err());
    }

    #[test]
    fn resolve_by_header_name() {
        let cols = vec!["State".to_string(), "Amount".to_string(), "Year".to_string()];
        assert_eq!(resolve_column("Amount", &cols).unwrap(), "Amount");
    }

    #[test]
    fn resolve_by_letter() {
        let cols = vec!["State".to_string(), "Amount".to_string(), "Year".to_string()];
        assert_eq!(resolve_column("B", &cols).unwrap(), "Amount");
    }

    #[test]
    fn resolve_by_letter_lowercase() {
        let cols = vec!["State".to_string(), "Amount".to_string(), "Year".to_string()];
        assert_eq!(resolve_column("b", &cols).unwrap(), "Amount");
    }

    #[test]
    fn resolve_header_takes_priority_over_letter() {
        let cols = vec!["A".to_string(), "B".to_string()];
        assert_eq!(resolve_column("A", &cols).unwrap(), "A");
    }

    #[test]
    fn resolve_case_insensitive_header() {
        let cols = vec!["State".to_string(), "Amount".to_string()];
        assert_eq!(resolve_column("state", &cols).unwrap(), "State");
    }

    #[test]
    fn resolve_unknown_column_is_err() {
        let cols = vec!["State".to_string(), "Amount".to_string()];
        let err = resolve_column("Foo", &cols).unwrap_err();
        assert!(err.contains("not found"), "error was: {}", err);
    }

    #[test]
    fn resolve_letter_out_of_range_is_err() {
        let cols = vec!["State".to_string()];
        let err = resolve_column("C", &cols).unwrap_err();
        assert!(err.contains("not found"), "error was: {}", err);
    }

    #[test]
    fn resolve_multiple_columns() {
        let cols = vec!["State".to_string(), "Amount".to_string(), "Year".to_string()];
        let resolved = resolve_columns(&["A".to_string(), "Year".to_string()], &cols).unwrap();
        assert_eq!(resolved, vec!["State", "Year"]);
    }
}
