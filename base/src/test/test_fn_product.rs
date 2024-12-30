#![allow(clippy::unwrap_used)]

use crate::test::util::new_empty_model;

#[test]
fn test_fn_product_arguments() {
    let mut model = new_empty_model();

    // Incorrect number of arguments
    model._set("A1", "=PRODUCT()");

    model.evaluate();
    // Error (Incorrect number of arguments)
    assert_eq!(model._get_text("A1"), *"#ERROR!");
}


#[test]
fn test_fn_product() {
    let mut model = new_empty_model();
    // Text converted to a number
    model._set("A1", r#"=PRODUCT("10")"#);
    model._set("A2", r#"=PRODUCT("1e2")"#);

    // Text in range not converted to a number
    model._set("A3", r#"=PRODUCT(B3:D3)"#);
    model._set("B3", r#""100"#);

    // Invalid text causes #VALUE!
    model._set("A4", r#"=PRODUCT("a")"#);

    // Implicit intersection not relevant for PRODUCT
    model._set("A5", r#"=PRODUCT(C1:C10)"#);
    model._set("C1", "2");
    model._set("C10", "3");

    // Boolean values are 1/0
    model._set("A6", r#"=PRODUCT(TRUE)"#);
    model._set("A7", r#"=PRODUCT(FALSE)"#);

    model.evaluate();

    assert_eq!(model._get_text("A1"), *"10");
    assert_eq!(model._get_text("A2"), *"100");
    assert_eq!(model._get_text("A3"), *"0");
    assert_eq!(model._get_text("A4"), *"#VALUE!");
    assert_eq!(model._get_text("A5"), *"6");
    assert_eq!(model._get_text("A6"), *"1");
    assert_eq!(model._get_text("A7"), *"0");
}
