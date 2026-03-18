use anyhow::Result;
use polars::prelude::*;

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

/// Check if a polars DataType is numeric.
fn is_numeric_dtype(dtype: &DataType) -> bool {
    matches!(
        dtype,
        DataType::Int8
            | DataType::Int16
            | DataType::Int32
            | DataType::Int64
            | DataType::UInt8
            | DataType::UInt16
            | DataType::UInt32
            | DataType::UInt64
            | DataType::Float32
            | DataType::Float64
    )
}

/// Build a boolean mask for a single filter expression against a DataFrame.
fn build_filter_mask(df: &DataFrame, expr: &FilterExpr) -> Result<BooleanChunked> {
    let col = df.column(&expr.column).map_err(|e| anyhow::anyhow!("{}", e))?;
    let series = col.as_materialized_series();
    let dtype = series.dtype();

    match &expr.op {
        FilterOp::Eq => {
            if is_numeric_dtype(dtype) {
                if let Ok(n) = expr.value.parse::<f64>() {
                    let s = series.cast(&DataType::Float64)?;
                    return Ok(s.f64()?.equal(n));
                }
            }
            let s = series.cast(&DataType::String)?;
            Ok(s.str()?.equal(expr.value.as_str()))
        }
        FilterOp::NotEq => {
            if is_numeric_dtype(dtype) {
                if let Ok(n) = expr.value.parse::<f64>() {
                    let s = series.cast(&DataType::Float64)?;
                    return Ok(s.f64()?.not_equal(n));
                }
            }
            let s = series.cast(&DataType::String)?;
            Ok(s.str()?.not_equal(expr.value.as_str()))
        }
        FilterOp::Gt => {
            let n = parse_numeric_value(&expr.value, ">")?;
            let s = series.cast(&DataType::Float64)?;
            Ok(s.f64()?.gt(n))
        }
        FilterOp::Lt => {
            let n = parse_numeric_value(&expr.value, "<")?;
            let s = series.cast(&DataType::Float64)?;
            Ok(s.f64()?.lt(n))
        }
        FilterOp::Gte => {
            let n = parse_numeric_value(&expr.value, ">=")?;
            let s = series.cast(&DataType::Float64)?;
            Ok(s.f64()?.gt_eq(n))
        }
        FilterOp::Lte => {
            let n = parse_numeric_value(&expr.value, "<=")?;
            let s = series.cast(&DataType::Float64)?;
            Ok(s.f64()?.lt_eq(n))
        }
        FilterOp::Contains => {
            let s = series.cast(&DataType::String)?;
            let ca = s.str()?;
            let pat = expr.value.to_lowercase();
            let mask: BooleanChunked = ca.into_iter()
                .map(|opt_s| opt_s.map(|s| s.to_lowercase().contains(&pat)).unwrap_or(false))
                .collect();
            Ok(mask)
        }
        FilterOp::NotContains => {
            let s = series.cast(&DataType::String)?;
            let ca = s.str()?;
            let pat = expr.value.to_lowercase();
            let mask: BooleanChunked = ca.into_iter()
                .map(|opt_s| opt_s.map(|s| !s.to_lowercase().contains(&pat)).unwrap_or(true))
                .collect();
            Ok(mask)
        }
    }
}

fn parse_numeric_value(value: &str, op: &str) -> Result<f64> {
    value
        .parse::<f64>()
        .map_err(|_| anyhow::anyhow!("'{}' requires numeric value, got '{}'", op, value))
}

/// Apply a list of filter expressions to a DataFrame (AND logic).
/// An empty list returns the DataFrame unchanged.
pub fn apply_filters(df: &DataFrame, exprs: &[FilterExpr]) -> Result<DataFrame> {
    let mut result = df.clone();
    for expr in exprs {
        let mask = build_filter_mask(&result, expr)?;
        result = result.filter(&mask)?;
    }
    Ok(result)
}

/// Options for the filter pipeline.
pub struct FilterOptions {
    pub filters: Vec<FilterExpr>,
    pub cols: Option<Vec<String>>,
    pub sort: Option<SortSpec>,
    pub limit: Option<usize>,
    pub head: Option<usize>,
    pub tail: Option<usize>,
}

/// Apply a sort specification to a DataFrame.
pub fn apply_sort(df: &DataFrame, spec: &SortSpec) -> Result<DataFrame> {
    let opts = SortMultipleOptions::default()
        .with_order_descending(spec.descending);
    Ok(df.sort([&spec.column], opts)?)
}

/// Run the full filter pipeline: head/tail → resolve & filter → sort → limit → select columns.
pub fn filter_pipeline(df: DataFrame, opts: &FilterOptions) -> Result<DataFrame> {
    let df_columns: Vec<String> = df
        .get_column_names()
        .iter()
        .map(|s| s.to_string())
        .collect();

    // 1. Pre-filter window: head or tail
    let df = if let Some(n) = opts.head {
        df.head(Some(n))
    } else if let Some(n) = opts.tail {
        df.tail(Some(n))
    } else {
        df
    };

    // 2. Resolve column names in filter expressions and apply filters
    let resolved_filters: Vec<FilterExpr> = opts
        .filters
        .iter()
        .map(|f| {
            let resolved_col = resolve_column(&f.column, &df_columns)?;
            Ok(FilterExpr {
                column: resolved_col,
                op: f.op.clone(),
                value: f.value.clone(),
            })
        })
        .collect::<Result<Vec<_>, String>>()
        .map_err(|e| anyhow::anyhow!("{}", e))?;

    let df = apply_filters(&df, &resolved_filters)?;

    // 3. Sort
    let df = if let Some(ref spec) = opts.sort {
        let resolved_col = resolve_column(&spec.column, &df_columns)
            .map_err(|e| anyhow::anyhow!("{}", e))?;
        let resolved_spec = SortSpec {
            column: resolved_col,
            descending: spec.descending,
        };
        apply_sort(&df, &resolved_spec)?
    } else {
        df
    };

    // 4. Limit (after filtering and sorting)
    let df = if let Some(n) = opts.limit {
        df.head(Some(n))
    } else {
        df
    };

    // 5. Select columns
    let df = if let Some(ref col_specs) = opts.cols {
        let resolved_cols = resolve_columns(col_specs, &df_columns)
            .map_err(|e| anyhow::anyhow!("{}", e))?;
        let col_refs: Vec<&str> = resolved_cols.iter().map(|s| s.as_str()).collect();
        df.select(col_refs)?
    } else {
        df
    };

    Ok(df)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn make_test_df() -> DataFrame {
        DataFrame::new(vec![
            Column::new("State".into(), &["CA", "NY", "CA", "TX", "NY"]),
            Column::new("City".into(), &["LA", "NYC", "SF", "Houston", "Albany"]),
            Column::new("Amount".into(), &[1500i64, 2000, 800, 1200, 500]),
            Column::new("Year".into(), &[2023i64, 2023, 2024, 2024, 2023]),
            Column::new("Status".into(), &["Active", "Active", "Draft", "Active", "Draft"]),
        ])
        .unwrap()
    }

    #[test]
    fn filter_eq_string() {
        let df = make_test_df();
        let expr = parse_filter_expr("State=CA").unwrap();
        let result = apply_filters(&df, &[expr]).unwrap();
        assert_eq!(result.height(), 2);
    }

    #[test]
    fn filter_eq_numeric() {
        let df = make_test_df();
        let expr = parse_filter_expr("Amount=1500").unwrap();
        let result = apply_filters(&df, &[expr]).unwrap();
        assert_eq!(result.height(), 1);
    }

    #[test]
    fn filter_not_eq() {
        let df = make_test_df();
        let expr = parse_filter_expr("Status!=Draft").unwrap();
        let result = apply_filters(&df, &[expr]).unwrap();
        assert_eq!(result.height(), 3);
    }

    #[test]
    fn filter_gt() {
        let df = make_test_df();
        let expr = parse_filter_expr("Amount>1000").unwrap();
        let result = apply_filters(&df, &[expr]).unwrap();
        assert_eq!(result.height(), 3);
    }

    #[test]
    fn filter_lt() {
        let df = make_test_df();
        let expr = parse_filter_expr("Amount<1000").unwrap();
        let result = apply_filters(&df, &[expr]).unwrap();
        assert_eq!(result.height(), 2);
    }

    #[test]
    fn filter_gte() {
        let df = make_test_df();
        let expr = parse_filter_expr("Amount>=1500").unwrap();
        let result = apply_filters(&df, &[expr]).unwrap();
        assert_eq!(result.height(), 2);
    }

    #[test]
    fn filter_lte() {
        let df = make_test_df();
        let expr = parse_filter_expr("Amount<=800").unwrap();
        let result = apply_filters(&df, &[expr]).unwrap();
        assert_eq!(result.height(), 2);
    }

    #[test]
    fn filter_contains() {
        let df = make_test_df();
        let expr = parse_filter_expr("City~ou").unwrap();
        let result = apply_filters(&df, &[expr]).unwrap();
        assert_eq!(result.height(), 1);
    }

    #[test]
    fn filter_contains_case_insensitive() {
        let df = make_test_df();
        let expr = parse_filter_expr("City~HOUSTON").unwrap();
        let result = apply_filters(&df, &[expr]).unwrap();
        assert_eq!(result.height(), 1);
    }

    #[test]
    fn filter_not_contains() {
        let df = make_test_df();
        let expr = parse_filter_expr("Status!~raft").unwrap();
        let result = apply_filters(&df, &[expr]).unwrap();
        assert_eq!(result.height(), 3);
    }

    #[test]
    fn filter_multiple_and() {
        let df = make_test_df();
        let e1 = parse_filter_expr("State=CA").unwrap();
        let e2 = parse_filter_expr("Amount>1000").unwrap();
        let result = apply_filters(&df, &[e1, e2]).unwrap();
        assert_eq!(result.height(), 1);
    }

    #[test]
    fn filter_no_matches_returns_empty() {
        let df = make_test_df();
        let expr = parse_filter_expr("State=ZZ").unwrap();
        let result = apply_filters(&df, &[expr]).unwrap();
        assert_eq!(result.height(), 0);
    }

    #[test]
    fn filter_empty_exprs_returns_all() {
        let df = make_test_df();
        let result = apply_filters(&df, &[]).unwrap();
        assert_eq!(result.height(), 5);
    }

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

    #[test]
    fn sort_ascending() {
        let df = make_test_df();
        let spec = parse_sort_spec("Amount:asc").unwrap();
        let result = apply_sort(&df, &spec).unwrap();
        let col = result.column("Amount").unwrap().as_materialized_series();
        let amounts = col.i64().unwrap();
        assert_eq!(amounts.get(0), Some(500));
        assert_eq!(amounts.get(4), Some(2000));
    }

    #[test]
    fn sort_descending() {
        let df = make_test_df();
        let spec = parse_sort_spec("Amount:desc").unwrap();
        let result = apply_sort(&df, &spec).unwrap();
        let col = result.column("Amount").unwrap().as_materialized_series();
        let amounts = col.i64().unwrap();
        assert_eq!(amounts.get(0), Some(2000));
        assert_eq!(amounts.get(4), Some(500));
    }

    #[test]
    fn pipeline_full() {
        let df = make_test_df();
        let opts = FilterOptions {
            filters: vec![parse_filter_expr("Amount>500").unwrap()],
            cols: Some(vec!["State".to_string(), "Amount".to_string()]),
            sort: Some(parse_sort_spec("Amount:desc").unwrap()),
            limit: Some(2),
            head: None,
            tail: None,
        };
        let result = filter_pipeline(df, &opts).unwrap();
        assert_eq!(result.height(), 2);
        assert_eq!(result.width(), 2);
        let col = result.column("Amount").unwrap().as_materialized_series();
        let amounts = col.i64().unwrap();
        assert_eq!(amounts.get(0), Some(2000));
        assert_eq!(amounts.get(1), Some(1500));
    }

    #[test]
    fn pipeline_head_before_filter() {
        let df = make_test_df(); // 5 rows: CA/LA, NY/NYC, CA/SF, TX/Houston, NY/Albany
        let opts = FilterOptions {
            filters: vec![parse_filter_expr("State=NY").unwrap()],
            cols: None,
            sort: None,
            limit: None,
            head: Some(3), // Take first 3 rows before filtering
            tail: None,
        };
        let result = filter_pipeline(df, &opts).unwrap();
        // First 3 rows: CA/LA, NY/NYC, CA/SF → only NY/NYC matches
        assert_eq!(result.height(), 1);
    }

    #[test]
    fn pipeline_tail_before_filter() {
        let df = make_test_df(); // 5 rows
        let opts = FilterOptions {
            filters: vec![parse_filter_expr("State=CA").unwrap()],
            cols: None,
            sort: None,
            limit: None,
            head: None,
            tail: Some(3), // Last 3 rows before filtering
        };
        let result = filter_pipeline(df, &opts).unwrap();
        // Last 3 rows: CA/SF, TX/Houston, NY/Albany → only CA/SF matches
        assert_eq!(result.height(), 1);
    }

    #[test]
    fn pipeline_no_options_returns_all() {
        let df = make_test_df();
        let opts = FilterOptions {
            filters: vec![],
            cols: None,
            sort: None,
            limit: None,
            head: None,
            tail: None,
        };
        let result = filter_pipeline(df, &opts).unwrap();
        assert_eq!(result.height(), 5);
        assert_eq!(result.width(), 5);
    }

    #[test]
    fn pipeline_cols_by_letter() {
        let df = make_test_df();
        let opts = FilterOptions {
            filters: vec![],
            cols: Some(vec!["A".to_string(), "C".to_string()]),
            sort: None,
            limit: None,
            head: None,
            tail: None,
        };
        let result = filter_pipeline(df, &opts).unwrap();
        assert_eq!(result.width(), 2);
        let names: Vec<String> = result.get_column_names().iter().map(|s| s.to_string()).collect();
        assert_eq!(names, vec!["State", "Amount"]);
    }

    #[test]
    fn pipeline_limit_after_filter() {
        let df = make_test_df();
        let opts = FilterOptions {
            filters: vec![parse_filter_expr("Status=Active").unwrap()],
            cols: None,
            sort: None,
            limit: Some(2),
            head: None,
            tail: None,
        };
        let result = filter_pipeline(df, &opts).unwrap();
        assert_eq!(result.height(), 2); // 3 Active rows, limited to 2
    }
}
