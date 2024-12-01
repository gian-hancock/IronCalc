use std::collections::HashMap;

use crate::{
    calc_result::CalcResult, expressions::{parser::Node, types::CellReferenceIndex}, language::Language, model::{CellState, Model}, number_format::to_precision, types::Workbook
};

impl Model {
    // DELTA(number1, [number2])
    pub(crate) fn fn_delta(
        workbook: &Workbook,
        cells: &mut HashMap<(u32, i32, i32), CellState>,
        parsed_formulas: &Vec<Vec<Node>>,
        language: &Language,
        args: &[Node], cell: CellReferenceIndex) -> CalcResult {
        let arg_count = args.len();
        if !(1..=2).contains(&arg_count) {
            return CalcResult::new_args_number_error(cell);
        }
        let number1 = match Model::get_number_no_bools(workbook, cells, parsed_formulas, language, &args[0], cell) {
            Ok(f) => f,
            Err(error) => return error,
        };
        let number2 = if arg_count > 1 {
            match Model::get_number_no_bools(workbook, cells, parsed_formulas, language, &args[1], cell) {
                Ok(f) => f,
                Err(error) => return error,
            }
        } else {
            0.0
        };

        if to_precision(number1, 16) == to_precision(number2, 16) {
            CalcResult::Number(1.0)
        } else {
            CalcResult::Number(0.0)
        }
    }

    // GESTEP(number, [step])
    pub(crate) fn fn_gestep(
        workbook: &Workbook,
        cells: &mut HashMap<(u32, i32, i32), CellState>,
        parsed_formulas: &Vec<Vec<Node>>,
        language: &Language,
        args: &[Node], cell: CellReferenceIndex) -> CalcResult {
        let arg_count = args.len();
        if !(1..=2).contains(&arg_count) {
            return CalcResult::new_args_number_error(cell);
        }
        let number = match Model::get_number_no_bools(workbook, cells, parsed_formulas, language, &args[0], cell) {
            Ok(f) => f,
            Err(error) => return error,
        };
        let step = if arg_count > 1 {
            match Model::get_number_no_bools(workbook, cells, parsed_formulas, language, &args[1], cell) {
                Ok(f) => f,
                Err(error) => return error,
            }
        } else {
            0.0
        };
        if to_precision(number, 16) >= to_precision(step, 16) {
            CalcResult::Number(1.0)
        } else {
            CalcResult::Number(0.0)
        }
    }
}
