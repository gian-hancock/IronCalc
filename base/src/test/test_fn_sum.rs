#![allow(clippy::unwrap_used)]

use crate::test::util::new_empty_model;

#[test]
fn test_fn_sum_arguments() {
    let mut model = new_empty_model();
    model._set("A1", "=SUM()");
    model._set("A2", "=SUM(1, 2, 3)");
    model._set("A3", "=SUM(1, )");
    model._set("A4", "=SUM(1,   , 3)");

    model.evaluate();

    assert_eq!(model._get_text("A1"), *"#ERROR!");
    assert_eq!(model._get_text("A2"), *"6");
    assert_eq!(model._get_text("A3"), *"1");
    assert_eq!(model._get_text("A4"), *"4");
}

#[test]
fn test_fn_sum() {
    let mut model = new_empty_model();
    // Text converted to a number
    model._set("A1", r#"=SUM("1")"#);
    model._set("A2", r#"=SUM("1e2")"#);

    // Text in range not converted to a number
    model._set("A3", r#"=SUM(B3:D3)"#);
    model._set("B3", r#"="100""#);

    // Invalid text causes #VALUE!
    model._set("A4", r#"=SUM("a")"#);

    // Implicit intersection not relevant for SUM
    model._set("A5", r#"=SUM(C1:C10)"#);
    model._set("C1", "1");
    model._set("C10", "1");

    // Boolean values are 1/0
    model._set("A6", r#"=SUM(TRUE)"#);
    model._set("A7", r#"=SUM(FALSE)"#);

    model.evaluate();

    assert_eq!(model._get_text("A1"), *"1");
    assert_eq!(model._get_text("A2"), *"100");
    assert_eq!(model._get_text("A3"), *"0");
    assert_eq!(model._get_text("A4"), *"#VALUE!");
    assert_eq!(model._get_text("A5"), *"2");
    assert_eq!(model._get_text("A6"), *"1");
    assert_eq!(model._get_text("A7"), *"0");
}

// WQ: Move to different file
#[test]
fn test_fn_product_arguments() {
    let mut model = new_empty_model();
    // WQ:
    model._set("A1", "=PRODUCT()");
    model._set("A2", "=PRODUCT(1, 2, 3)");
    model._set("A3", "=PRODUCT(1, )");
    model._set("A4", "=PRODUCT(1,   , 3)");

    model.evaluate();

    assert_eq!(model._get_text("A1"), *"#ERROR!");
    assert_eq!(model._get_text("A2"), *"6");
    assert_eq!(model._get_text("A3"), *"1");
    assert_eq!(model._get_text("A4"), *"3");
}

// WQ: Move to different file
#[test]
fn test_fn_product() {
    let mut model = new_empty_model();
    // Text converted to a number
    model._set("A1", r#"=PRODUCT("1")"#);
    model._set("A2", r#"=PRODUCT("1e2")"#);

    // Text in range not converted to a number
    model._set("A3", r#"=PRODUCT(B3:D3)"#);
    model._set("B3", r#"="100""#);

    // Invalid text causes #VALUE!
    model._set("A4", r#"=PRODUCT("a")"#);

    // Implicit intersection not relevant for PRODUCT
    model._set("A5", r#"=PRODUCT(C1:C10)"#);
    model._set("C1", "1");
    model._set("C10", "1");

    // Boolean values are 1/0
    model._set("A6", r#"=PRODUCT(TRUE)"#);
    model._set("A7", r#"=PRODUCT(FALSE)"#);

    model.evaluate();

    assert_eq!(model._get_text("A1"), *"1");
    assert_eq!(model._get_text("A2"), *"100");
    assert_eq!(model._get_text("A3"), *"0");
    assert_eq!(model._get_text("A4"), *"#VALUE!");
    assert_eq!(model._get_text("A5"), *"1");
    assert_eq!(model._get_text("A6"), *"1");
    assert_eq!(model._get_text("A7"), *"0");
}
