#![allow(clippy::unwrap_used)]

use crate::{
    constants::{LAST_COLUMN, LAST_ROW},
    test::util::new_empty_model, Model,
};

#[test]
fn test_empty_model() {
    let mut model = new_empty_model();
    assert_eq!(Model::get_frozen_rows_count(&model.workbook, 0), Ok(0));
    assert_eq!(Model::get_frozen_columns_count(&model.workbook, 0), Ok(0));

    let e = Model::set_frozen_rows(&mut model.workbook, 0, 3);
    assert!(e.is_ok());
    assert_eq!(Model::get_frozen_rows_count(&model.workbook, 0), Ok(3));
    assert_eq!(Model::get_frozen_columns_count(&model.workbook, 0), Ok(0));

    let e = Model::set_frozen_columns(&mut model.workbook, 0, 53);
    assert!(e.is_ok());
    assert_eq!(Model::get_frozen_rows_count(&model.workbook, 0), Ok(3));
    assert_eq!(Model::get_frozen_columns_count(&model.workbook, 0), Ok(53));

    // Set them back to zero
    let e = Model::set_frozen_rows(&mut model.workbook, 0, 0);
    assert!(e.is_ok());
    let e = Model::set_frozen_columns(&mut model.workbook, 0, 0);
    assert!(e.is_ok());
    assert_eq!(Model::get_frozen_rows_count(&model.workbook, 0), Ok(0));
    assert_eq!(Model::get_frozen_columns_count(&model.workbook, 0), Ok(0));
}

#[test]
fn test_invalid_sheet() {
    let mut model = new_empty_model();
    assert_eq!(
        Model::get_frozen_rows_count(&model.workbook, 1),
        Err("Invalid sheet".to_string())
    );
    assert_eq!(
        Model::get_frozen_columns_count(&model.workbook, 3),
        Err("Invalid sheet".to_string())
    );

    assert_eq!(
        Model::set_frozen_rows(&mut model.workbook, 3, 3),
        Err("Invalid sheet".to_string())
    );
    assert_eq!(
        Model::set_frozen_columns(&mut model.workbook, 3, 5),
        Err("Invalid sheet".to_string())
    );
}

#[test]
fn test_invalid_rows_columns() {
    let mut model = new_empty_model();

    assert_eq!(
        Model::set_frozen_rows(&mut model.workbook, 0, -3),
        Err("Frozen rows cannot be negative".to_string())
    );
    assert_eq!(
        Model::set_frozen_columns(&mut model.workbook, 0, -5),
        Err("Frozen columns cannot be negative".to_string())
    );

    assert_eq!(
        Model::set_frozen_rows(&mut model.workbook, 0, LAST_ROW),
        Err("Too many rows".to_string())
    );
    assert_eq!(
        Model::set_frozen_columns(&mut model.workbook, 0, LAST_COLUMN),
        Err("Too many columns".to_string())
    );
}
