use std::collections::HashMap;

use crate::{
    calc_result::CalcResult, expressions::{parser::Node, token::Error, types::CellReferenceIndex}, language::Language, model::{Model, ParsedDefinedName}, types::Workbook
};

use super::CellState;

impl Model {
    pub(crate) fn fn_isnumber(
        workbook: &Workbook,
        cells: &mut HashMap<(u32, i32, i32), CellState>,
        parsed_formulas: &Vec<Vec<Node>>,
        language: &Language,
        args: &[Node], cell: CellReferenceIndex) -> CalcResult {
        if args.len() == 1 {
            match Model::evaluate_node_in_context(workbook, cells, parsed_formulas, language, &args[0], cell) {
                CalcResult::Number(_) => return CalcResult::Boolean(true),
                _ => {
                    return CalcResult::Boolean(false);
                }
            };
        }
        CalcResult::new_args_number_error(cell)
    }
    pub(crate) fn fn_istext(
        workbook: &Workbook,
        cells: &mut HashMap<(u32, i32, i32), CellState>,
        parsed_formulas: &Vec<Vec<Node>>,
        language: &Language,
        args: &[Node], cell: CellReferenceIndex) -> CalcResult {
        if args.len() == 1 {
            match Model::evaluate_node_in_context(workbook, cells, parsed_formulas, language, &args[0], cell) {
                CalcResult::String(_) => return CalcResult::Boolean(true),
                _ => {
                    return CalcResult::Boolean(false);
                }
            };
        }
        CalcResult::new_args_number_error(cell)
    }
    pub(crate) fn fn_isnontext(
        workbook: &Workbook,
        cells: &mut HashMap<(u32, i32, i32), CellState>,
        parsed_formulas: &Vec<Vec<Node>>,
        language: &Language,
        args: &[Node], cell: CellReferenceIndex) -> CalcResult {
        if args.len() == 1 {
            match Model::evaluate_node_in_context(workbook, cells, parsed_formulas, language, &args[0], cell) {
                CalcResult::String(_) => return CalcResult::Boolean(false),
                _ => {
                    return CalcResult::Boolean(true);
                }
            };
        }
        CalcResult::new_args_number_error(cell)
    }
    pub(crate) fn fn_islogical(
        workbook: &Workbook,
        cells: &mut HashMap<(u32, i32, i32), CellState>,
        parsed_formulas: &Vec<Vec<Node>>,
        language: &Language,
        args: &[Node], cell: CellReferenceIndex) -> CalcResult {
        if args.len() == 1 {
            match Model::evaluate_node_in_context(workbook, cells, parsed_formulas, language, &args[0], cell) {
                CalcResult::Boolean(_) => return CalcResult::Boolean(true),
                _ => {
                    return CalcResult::Boolean(false);
                }
            };
        }
        CalcResult::new_args_number_error(cell)
    }
    pub(crate) fn fn_isblank(
        workbook: &Workbook,
        cells: &mut HashMap<(u32, i32, i32), CellState>,
        parsed_formulas: &Vec<Vec<Node>>,
        language: &Language,
        args: &[Node], cell: CellReferenceIndex) -> CalcResult {
        if args.len() == 1 {
            match Model::evaluate_node_in_context(workbook, cells, parsed_formulas, language, &args[0], cell) {
                CalcResult::EmptyCell => return CalcResult::Boolean(true),
                _ => {
                    return CalcResult::Boolean(false);
                }
            };
        }
        CalcResult::new_args_number_error(cell)
    }
    pub(crate) fn fn_iserror(
        workbook: &Workbook,
        cells: &mut HashMap<(u32, i32, i32), CellState>,
        parsed_formulas: &Vec<Vec<Node>>,
        language: &Language,
        args: &[Node], cell: CellReferenceIndex) -> CalcResult {
        if args.len() == 1 {
            match Model::evaluate_node_in_context(workbook, cells, parsed_formulas, language, &args[0], cell) {
                CalcResult::Error { .. } => return CalcResult::Boolean(true),
                _ => {
                    return CalcResult::Boolean(false);
                }
            };
        }
        CalcResult::new_args_number_error(cell)
    }
    pub(crate) fn fn_iserr(
        workbook: &Workbook,
        cells: &mut HashMap<(u32, i32, i32), CellState>,
        parsed_formulas: &Vec<Vec<Node>>,
        language: &Language,
        args: &[Node], cell: CellReferenceIndex) -> CalcResult {
        if args.len() == 1 {
            match Model::evaluate_node_in_context(workbook, cells, parsed_formulas, language, &args[0], cell) {
                CalcResult::Error { error, .. } => {
                    if Error::NA == error {
                        return CalcResult::Boolean(false);
                    } else {
                        return CalcResult::Boolean(true);
                    }
                }
                _ => {
                    return CalcResult::Boolean(false);
                }
            };
        }
        CalcResult::new_args_number_error(cell)
    }
    pub(crate) fn fn_isna(
        workbook: &Workbook,
        cells: &mut HashMap<(u32, i32, i32), CellState>,
        parsed_formulas: &Vec<Vec<Node>>,
        language: &Language,
        args: &[Node], cell: CellReferenceIndex) -> CalcResult {
        if args.len() == 1 {
            match Model::evaluate_node_in_context(workbook, cells, parsed_formulas, language, &args[0], cell) {
                CalcResult::Error { error, .. } => {
                    if error == Error::NA {
                        return CalcResult::Boolean(true);
                    } else {
                        return CalcResult::Boolean(false);
                    }
                }
                _ => {
                    return CalcResult::Boolean(false);
                }
            };
        }
        CalcResult::new_args_number_error(cell)
    }

    // Returns true if it is a reference or evaluates to a reference
    // But DOES NOT evaluate
    pub(crate) fn fn_isref(
        workbook: &Workbook,
        cells: &mut HashMap<(u32, i32, i32), CellState>,
        parsed_formulas: &Vec<Vec<Node>>,
        language: &Language,
        args: &[Node], cell: CellReferenceIndex) -> CalcResult {
        if args.len() != 1 {
            return CalcResult::new_args_number_error(cell);
        }
        match &args[0] {
            Node::ReferenceKind { .. } | Node::RangeKind { .. } | Node::OpRangeKind { .. } => {
                CalcResult::Boolean(true)
            }
            Node::FunctionKind { kind, args: _ } => CalcResult::Boolean(kind.returns_reference()),
            _ => CalcResult::Boolean(false),
        }
    }

    pub(crate) fn fn_isodd(
        workbook: &Workbook,
        cells: &mut HashMap<(u32, i32, i32), CellState>,
        parsed_formulas: &Vec<Vec<Node>>,
        language: &Language,
        args: &[Node], cell: CellReferenceIndex) -> CalcResult {
        if args.len() != 1 {
            return CalcResult::new_args_number_error(cell);
        }
        let value = match Model::get_number_no_bools(workbook, cells, parsed_formulas, language, &args[0], cell) {
            Ok(f) => f.abs().trunc() as i64,
            Err(s) => return s,
        };
        CalcResult::Boolean(value % 2 == 1)
    }

    pub(crate) fn fn_iseven(
        workbook: &Workbook,
        cells: &mut HashMap<(u32, i32, i32), CellState>,
        parsed_formulas: &Vec<Vec<Node>>,
        language: &Language,
        args: &[Node], cell: CellReferenceIndex) -> CalcResult {
        if args.len() != 1 {
            return CalcResult::new_args_number_error(cell);
        }
        let value = match Model::get_number_no_bools(workbook, cells, parsed_formulas, language, &args[0], cell) {
            Ok(f) => f.abs().trunc() as i64,
            Err(s) => return s,
        };
        CalcResult::Boolean(value % 2 == 0)
    }

    // ISFORMULA arg needs to be a reference or something that evaluates to a reference
    pub(crate) fn fn_isformula(
        workbook: &Workbook,
        cells: &mut HashMap<(u32, i32, i32), CellState>,
        parsed_formulas: &Vec<Vec<Node>>,
        language: &Language,
        args: &[Node], cell: CellReferenceIndex) -> CalcResult {
        if args.len() != 1 {
            return CalcResult::new_args_number_error(cell);
        }
        if let CalcResult::Range { left, right } = Model::evaluate_node_with_reference(workbook, cells, parsed_formulas, language, &args[0], cell)
        {
            if left.sheet != right.sheet {
                return CalcResult::Error {
                    error: Error::ERROR,
                    origin: cell,
                    message: "3D ranges not supported".to_string(),
                };
            }
            if left.row != right.row && left.column != right.column {
                // FIXME: Implicit intersection or dynamic arrays
                return CalcResult::Error {
                    error: Error::VALUE,
                    origin: cell,
                    message: "argument must be a reference to a single cell".to_string(),
                };
            }
            let is_formula = if let Ok(f) = Model::get_cell_formula(workbook, cells, parsed_formulas, language, left.sheet, left.row, left.column)
            {
                f.is_some()
            } else {
                false
            };
            CalcResult::Boolean(is_formula)
        } else {
            CalcResult::Error {
                error: Error::ERROR,
                origin: cell,
                message: "Argument must be a reference".to_string(),
            }
        }
    }

    pub(crate) fn fn_errortype(
        workbook: &Workbook,
        cells: &mut HashMap<(u32, i32, i32), CellState>,
        parsed_formulas: &Vec<Vec<Node>>,
        language: &Language,
        args: &[Node], cell: CellReferenceIndex) -> CalcResult {
        if args.len() != 1 {
            return CalcResult::new_args_number_error(cell);
        }
        match Model::evaluate_node_in_context(workbook, cells, parsed_formulas, language, &args[0], cell) {
            CalcResult::Error { error, .. } => {
                match error {
                    Error::NULL => CalcResult::Number(1.0),
                    Error::DIV => CalcResult::Number(2.0),
                    Error::VALUE => CalcResult::Number(3.0),
                    Error::REF => CalcResult::Number(4.0),
                    Error::NAME => CalcResult::Number(5.0),
                    Error::NUM => CalcResult::Number(6.0),
                    Error::NA => CalcResult::Number(7.0),
                    Error::SPILL => CalcResult::Number(9.0),
                    Error::CALC => CalcResult::Number(14.0),
                    // IronCalc specific
                    Error::ERROR => CalcResult::Number(101.0),
                    Error::NIMPL => CalcResult::Number(102.0),
                    Error::CIRC => CalcResult::Number(104.0),
                    // Missing from Excel
                    // #GETTING_DATA => 8
                    // #CONNECT => 10
                    // #BLOCKED => 11
                    // #UNKNOWN => 12
                    // #FIELD => 13
                    // #EXTERNAL => 19
                }
            }
            _ => CalcResult::Error {
                error: Error::NA,
                origin: cell,
                message: "Not an error".to_string(),
            },
        }
    }

    // Excel believes for some reason that TYPE(A1:A7) is an array formula
    // Although we evaluate the same as Excel we cannot, ATM import this from excel
    pub(crate) fn fn_type(
        workbook: &Workbook,
        cells: &mut HashMap<(u32, i32, i32), CellState>,
        parsed_formulas: &Vec<Vec<Node>>,
        language: &Language,
        args: &[Node], cell: CellReferenceIndex) -> CalcResult {
        if args.len() != 1 {
            return CalcResult::new_args_number_error(cell);
        }
        match Model::evaluate_node_in_context(workbook, cells, parsed_formulas, language, &args[0], cell) {
            CalcResult::String(_) => CalcResult::Number(2.0),
            CalcResult::Number(_) => CalcResult::Number(1.0),
            CalcResult::Boolean(_) => CalcResult::Number(4.0),
            CalcResult::Error { .. } => CalcResult::Number(16.0),
            CalcResult::Range { .. } => CalcResult::Number(64.0),
            CalcResult::EmptyCell => CalcResult::Number(1.0),
            CalcResult::EmptyArg => {
                // This cannot happen
                CalcResult::Number(1.0)
            }
        }
    }
    pub(crate) fn fn_sheet(
        workbook: &Workbook,
        cells: &mut HashMap<(u32, i32, i32), CellState>,
        parsed_formulas: &Vec<Vec<Node>>,
        language: &Language,
        args: &[Node], cell: CellReferenceIndex) -> CalcResult {
        let arg_count = args.len();
        if arg_count > 1 {
            return CalcResult::new_args_number_error(cell);
        }
        if arg_count == 0 {
            // Sheets are 0-indexed`
            return CalcResult::Number(cell.sheet as f64 + 1.0);
        }
        // The arg could be a defined name or a table
        let arg = &args[0];
        if let Node::VariableKind(name) = arg {
            // Let's see if it is a defined name
            if let Some(defined_name) = parsed_defined_names.get(&(None, name.to_lowercase()))
            {
                match defined_name {
                    ParsedDefinedName::CellReference(reference) => {
                        return CalcResult::Number(reference.sheet as f64 + 1.0)
                    }
                    ParsedDefinedName::RangeReference(range) => {
                        return CalcResult::Number(range.left.sheet as f64 + 1.0)
                    }
                    ParsedDefinedName::InvalidDefinedNameFormula => {
                        return CalcResult::Error {
                            error: Error::NA,
                            origin: cell,
                            message: "Invalid name".to_string(),
                        };
                    }
                }
            }
            // Now let's see if it is a table
            for (table_name, table) in workbook.tables {
                if table_name == name {
                    if let Some(sheet_index) = Model::get_sheet_index_by_name(workbook, &table.sheet_name) {
                        return CalcResult::Number(sheet_index as f64 + 1.0);
                    } else {
                        break;
                    }
                }
            }
        }
        // Now it should be the name of a sheet
        let sheet_name = match Model::get_string(workbook, cells, parsed_formulas, language, arg, cell) {
            Ok(s) => s,
            Err(e) => return e,
        };
        if let Some(sheet_index) = Model::get_sheet_index_by_name(workbook, &sheet_name) {
            return CalcResult::Number(sheet_index as f64 + 1.0);
        }
        CalcResult::Error {
            error: Error::NA,
            origin: cell,
            message: "Invalid name".to_string(),
        }
    }
}
