#![allow(clippy::unwrap_used)]

use crate::test::util::new_empty_model;

#[test]
fn fn_and_1_arg() {
    let mut model = new_empty_model();
    model._set("A1", "=AND(1)");
    model._set("A2", "=AND(0)");
    model.evaluate();
    assert_eq!(model._get_text("A1"), *"TRUE");
    assert_eq!(model._get_text("A2"), *"FALSE");
}

#[test]
fn fn_and_2_arg() {
    let mut model = new_empty_model();
    model._set("A1", "=AND(1, 0)");
    model._set("A2", "=AND(1, 1)");
    model.evaluate();
    assert_eq!(model._get_text("A1"), *"FALSE");
    assert_eq!(model._get_text("A2"), *"TRUE");
}

#[test]
fn fn_and_missing_args() {
    let mut model = new_empty_model();
    model._set("A1", "=AND()");
    model.evaluate();
    assert_eq!(model._get_text("A1"), *"#ERROR!");
}

#[test]
fn fn_and_empty_arg() {
    let mut model = new_empty_model();
    model._set("A1", "=AND(A2)");

    model.evaluate();

    // assert_eq!(model._get_text("A1"), *"0");
    assert_eq!(model._get_text("A2"), *"5");
    assert_eq!(model._get_text("A3"), *"0");
}