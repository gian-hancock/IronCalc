use std::collections::HashMap;

use crate::{
    calc_result::CalcResult, expressions::{
        parser::{parse_range, Node},
        token::Error,
        types::CellReferenceIndex,
    }, functions::Function, language::Language, model::Model, types::Workbook
};

use super::CellState;

/// Excel has a complicated way of filtering + hidden rows
/// As a first a approximation a table can either have filtered rows or hidden rows, but not both.
/// Internally hey both will be marked as hidden rows. Hidden rows
/// The behaviour is the same for SUBTOTAL(100s,) => ignore those
/// But changes for the SUBTOTAL(1-11, ) those ignore filtered but take hidden into account.
/// In Excel filters are non-dynamic. Once you apply filters in a table (say value in column 1 should > 20) they
/// stay that way, even if you change the value of the values in the table after the fact.
/// If you try to hide rows in a table with filtered rows they will behave as if filtered
/// // Also subtotals ignore subtotals
///
#[derive(PartialEq)]
enum SubTotalMode {
    Full,
    SkipHidden,
}

#[derive(PartialEq, Debug)]
pub enum CellTableStatus {
    Normal,
    Hidden,
    Filtered,
}

impl Model {
    fn get_table_for_cell(
        workbook: &Workbook,
        cells: &mut HashMap<(u32, i32, i32), CellState>,
        parsed_formulas: &Vec<Vec<Node>>,
        language: &Language,
        sheet_index: u32, row: i32, column: i32) -> bool {
        let worksheet = match workbook.worksheet(sheet_index) {
            Ok(ws) => ws,
            Err(_) => return false,
        };
        for table in workbook.tables.values() {
            if worksheet.name != table.sheet_name {
                continue;
            }
            // (column, row, column, row)
            if let Ok((column1, row1, column2, row2)) = parse_range(&table.reference) {
                if ((column >= column1) && (column <= column2)) && ((row >= row1) && (row <= row2))
                {
                    return table.has_filters;
                }
            }
        }
        false
    }

    fn cell_hidden_status(
        workbook: &Workbook,
        cells: &mut HashMap<(u32, i32, i32), CellState>,
        parsed_formulas: &Vec<Vec<Node>>,
        language: &Language,
        sheet_index: u32, row: i32, column: i32) -> CellTableStatus {
        let worksheet = workbook.worksheet(sheet_index).expect("");
        let mut hidden = false;
        for row_style in &worksheet.rows {
            if row_style.r == row {
                hidden = row_style.hidden;
                break;
            }
        }
        if !hidden {
            return CellTableStatus::Normal;
        }
        // The row is hidden we need to know if the table has filters
        if Model::get_table_for_cell(workbook, cells, parsed_formulas, language, sheet_index, row, column) {
            CellTableStatus::Filtered
        } else {
            CellTableStatus::Hidden
        }
    }

    // FIXME(TD): This is too much
    fn cell_is_subtotal(
        workbook: &Workbook,
        cells: &mut HashMap<(u32, i32, i32), CellState>,
        parsed_formulas: &Vec<Vec<Node>>,
        language: &Language,
        sheet_index: u32, row: i32, column: i32) -> bool {
        let row_data = match workbook.worksheets[sheet_index as usize]
            .sheet_data
            .get(&row)
        {
            Some(r) => r,
            None => return false,
        };
        let cell = match row_data.get(&column) {
            Some(c) => c,
            None => {
                return false;
            }
        };

        match cell.get_formula() {
            Some(f) => {
                let node = &parsed_formulas[sheet_index as usize][f as usize];
                matches!(
                    node,
                    Node::FunctionKind {
                        kind: Function::Subtotal,
                        args: _
                    }
                )
            }
            None => false,
        }
    }

    fn subtotal_get_values(
        workbook: &Workbook,
        cells: &mut HashMap<(u32, i32, i32), CellState>,
        parsed_formulas: &Vec<Vec<Node>>,
        language: &Language,
        args: &[Node],
        cell: CellReferenceIndex,
        mode: SubTotalMode,
    ) -> Result<Vec<f64>, CalcResult> {
        let mut result: Vec<f64> = Vec::new();
        for arg in args {
            match arg {
                Node::FunctionKind {
                    kind: Function::Subtotal,
                    args: _,
                } => {
                    // skip
                }
                _ => {
                    match Model::evaluate_node_with_reference(workbook, cells, parsed_formulas, language, arg, cell) {
                        CalcResult::String(_) | CalcResult::Boolean(_) => {
                            // Skip
                        }
                        CalcResult::Number(f) => result.push(f),
                        error @ CalcResult::Error { .. } => {
                            return Err(error);
                        }
                        CalcResult::Range { left, right } => {
                            if left.sheet != right.sheet {
                                return Err(CalcResult::new_error(
                                    Error::VALUE,
                                    cell,
                                    "Ranges are in different sheets".to_string(),
                                ));
                            }
                            // We are not expecting subtotal to have open ranges
                            let row1 = left.row;
                            let row2 = right.row;
                            let column1 = left.column;
                            let column2 = right.column;

                            for row in row1..=row2 {
                                let cell_status = Model::cell_hidden_status(workbook, cells, parsed_formulas, language, left.sheet, row, column1);
                                if cell_status == CellTableStatus::Filtered {
                                    continue;
                                }
                                if mode == SubTotalMode::SkipHidden
                                    && cell_status == CellTableStatus::Hidden
                                {
                                    continue;
                                }
                                for column in column1..=column2 {
                                    if Model::cell_is_subtotal(workbook, cells, parsed_formulas, language, left.sheet, row, column) {
                                        continue;
                                    }
                                    match Model::evaluate_cell(workbook, cells, parsed_formulas, language, CellReferenceIndex {
                                        sheet: left.sheet,
                                        row,
                                        column,
                                    }) {
                                        CalcResult::Number(value) => {
                                            result.push(value);
                                        }
                                        error @ CalcResult::Error { .. } => return Err(error),
                                        _ => {
                                            // We ignore booleans and strings
                                        }
                                    }
                                }
                            }
                        }
                        CalcResult::EmptyCell | CalcResult::EmptyArg => result.push(0.0),
                    }
                }
            }
        }
        Ok(result)
    }

    pub(crate) fn fn_subtotal(
        workbook: &Workbook,
        cells: &mut HashMap<(u32, i32, i32), CellState>,
        parsed_formulas: &Vec<Vec<Node>>,
        language: &Language,
        args: &[Node], cell: CellReferenceIndex) -> CalcResult {
        if args.len() < 2 {
            return CalcResult::new_args_number_error(cell);
        }
        let value = match Model::get_number(workbook, cells, parsed_formulas, language, &args[0], cell) {
            Ok(f) => f.trunc() as i32,
            Err(s) => return s,
        };
        match value {
            1 => Model::subtotal_average(workbook, cells, parsed_formulas, language, &args[1..], cell, SubTotalMode::Full),
            2 => Model::subtotal_count(workbook, cells, parsed_formulas, language, &args[1..], cell, SubTotalMode::Full),
            3 => Model::subtotal_counta(workbook, cells, parsed_formulas, language, &args[1..], cell, SubTotalMode::Full),
            4 => Model::subtotal_max(workbook, cells, parsed_formulas, language, &args[1..], cell, SubTotalMode::Full),
            5 => Model::subtotal_min(workbook, cells, parsed_formulas, language, &args[1..], cell, SubTotalMode::Full),
            6 => Model::subtotal_product(workbook, cells, parsed_formulas, language, &args[1..], cell, SubTotalMode::Full),
            7 => Model::subtotal_stdevs(workbook, cells, parsed_formulas, language, &args[1..], cell, SubTotalMode::Full),
            8 => Model::subtotal_stdevp(workbook, cells, parsed_formulas, language, &args[1..], cell, SubTotalMode::Full),
            9 => Model::subtotal_sum(workbook, cells, parsed_formulas, language, &args[1..], cell, SubTotalMode::Full),
            10 => Model::subtotal_vars(workbook, cells, parsed_formulas, language, &args[1..], cell, SubTotalMode::Full),
            11 => Model::subtotal_varp(workbook, cells, parsed_formulas, language, &args[1..], cell, SubTotalMode::Full),
            101 => Model::subtotal_average(workbook, cells, parsed_formulas, language, &args[1..], cell, SubTotalMode::SkipHidden),
            102 => Model::subtotal_count(workbook, cells, parsed_formulas, language, &args[1..], cell, SubTotalMode::SkipHidden),
            103 => Model::subtotal_counta(workbook, cells, parsed_formulas, language, &args[1..], cell, SubTotalMode::SkipHidden),
            104 => Model::subtotal_max(workbook, cells, parsed_formulas, language, &args[1..], cell, SubTotalMode::SkipHidden),
            105 => Model::subtotal_min(workbook, cells, parsed_formulas, language, &args[1..], cell, SubTotalMode::SkipHidden),
            106 => Model::subtotal_product(workbook, cells, parsed_formulas, language, &args[1..], cell, SubTotalMode::SkipHidden),
            107 => Model::subtotal_stdevs(workbook, cells, parsed_formulas, language, &args[1..], cell, SubTotalMode::SkipHidden),
            108 => Model::subtotal_stdevp(workbook, cells, parsed_formulas, language, &args[1..], cell, SubTotalMode::SkipHidden),
            109 => Model::subtotal_sum(workbook, cells, parsed_formulas, language, &args[1..], cell, SubTotalMode::SkipHidden),
            110 => Model::subtotal_vars(workbook, cells, parsed_formulas, language, &args[1..], cell, SubTotalMode::Full),
            111 => Model::subtotal_varp(workbook, cells, parsed_formulas, language, &args[1..], cell, SubTotalMode::Full),
            _ => CalcResult::new_error(
                Error::VALUE,
                cell,
                format!("Invalid value for SUBTOTAL: {value}"),
            ),
        }
    }

    fn subtotal_vars(
        workbook: &Workbook,
        cells: &mut HashMap<(u32, i32, i32), CellState>,
        parsed_formulas: &Vec<Vec<Node>>,
        language: &Language,
        args: &[Node],
        cell: CellReferenceIndex,
        mode: SubTotalMode,
    ) -> CalcResult {
        let values = match Model::subtotal_get_values(workbook, cells, parsed_formulas, language, args, cell, mode) {
            Ok(s) => s,
            Err(s) => return s,
        };
        let mut result = 0.0;
        let l = values.len();
        for value in &values {
            result += value;
        }
        if l < 2 {
            return CalcResult::Error {
                error: Error::DIV,
                origin: cell,
                message: "Division by 0!".to_string(),
            };
        }
        // average
        let average = result / (l as f64);
        let mut result = 0.0;
        for value in &values {
            result += (value - average).powi(2) / (l as f64 - 1.0)
        }

        CalcResult::Number(result)
    }

    fn subtotal_varp(
        workbook: &Workbook,
        cells: &mut HashMap<(u32, i32, i32), CellState>,
        parsed_formulas: &Vec<Vec<Node>>,
        language: &Language,
        args: &[Node],
        cell: CellReferenceIndex,
        mode: SubTotalMode,
    ) -> CalcResult {
        let values = match Model::subtotal_get_values(workbook, cells, parsed_formulas, language, args, cell, mode) {
            Ok(s) => s,
            Err(s) => return s,
        };
        let mut result = 0.0;
        let l = values.len();
        for value in &values {
            result += value;
        }
        if l == 0 {
            return CalcResult::Error {
                error: Error::DIV,
                origin: cell,
                message: "Division by 0!".to_string(),
            };
        }
        // average
        let average = result / (l as f64);
        let mut result = 0.0;
        for value in &values {
            result += (value - average).powi(2) / (l as f64)
        }
        CalcResult::Number(result)
    }

    fn subtotal_stdevs(
        workbook: &Workbook,
        cells: &mut HashMap<(u32, i32, i32), CellState>,
        parsed_formulas: &Vec<Vec<Node>>,
        language: &Language,
        args: &[Node],
        cell: CellReferenceIndex,
        mode: SubTotalMode,
    ) -> CalcResult {
        let values = match Model::subtotal_get_values(workbook, cells, parsed_formulas, language, args, cell, mode) {
            Ok(s) => s,
            Err(s) => return s,
        };
        let mut result = 0.0;
        let l = values.len();
        for value in &values {
            result += value;
        }
        if l < 2 {
            return CalcResult::Error {
                error: Error::DIV,
                origin: cell,
                message: "Division by 0!".to_string(),
            };
        }
        // average
        let average = result / (l as f64);
        let mut result = 0.0;
        for value in &values {
            result += (value - average).powi(2) / (l as f64 - 1.0)
        }

        CalcResult::Number(result.sqrt())
    }

    fn subtotal_stdevp(
        workbook: &Workbook,
        cells: &mut HashMap<(u32, i32, i32), CellState>,
        parsed_formulas: &Vec<Vec<Node>>,
        language: &Language,
        args: &[Node],
        cell: CellReferenceIndex,
        mode: SubTotalMode,
    ) -> CalcResult {
        let values = match Model::subtotal_get_values(workbook, cells, parsed_formulas, language, args, cell, mode) {
            Ok(s) => s,
            Err(s) => return s,
        };
        let mut result = 0.0;
        let l = values.len();
        for value in &values {
            result += value;
        }
        if l == 0 {
            return CalcResult::Error {
                error: Error::DIV,
                origin: cell,
                message: "Division by 0!".to_string(),
            };
        }
        // average
        let average = result / (l as f64);
        let mut result = 0.0;
        for value in &values {
            result += (value - average).powi(2) / (l as f64)
        }
        CalcResult::Number(result.sqrt())
    }

    fn subtotal_counta(
        workbook: &Workbook,
        cells: &mut HashMap<(u32, i32, i32), CellState>,
        parsed_formulas: &Vec<Vec<Node>>,
        language: &Language,
        args: &[Node],
        cell: CellReferenceIndex,
        mode: SubTotalMode,
    ) -> CalcResult {
        let mut counta = 0;
        for arg in args {
            match arg {
                Node::FunctionKind {
                    kind: Function::Subtotal,
                    args: _,
                } => {
                    // skip
                }
                _ => {
                    match Model::evaluate_node_with_reference(workbook, cells, parsed_formulas, language, arg, cell) {
                        CalcResult::EmptyCell | CalcResult::EmptyArg => {
                            // skip
                        }
                        CalcResult::Range { left, right } => {
                            if left.sheet != right.sheet {
                                return CalcResult::new_error(
                                    Error::VALUE,
                                    cell,
                                    "Ranges are in different sheets".to_string(),
                                );
                            }
                            // We are not expecting subtotal to have open ranges
                            let row1 = left.row;
                            let row2 = right.row;
                            let column1 = left.column;
                            let column2 = right.column;

                            for row in row1..=row2 {
                                let cell_status = Model::cell_hidden_status(workbook, cells, parsed_formulas, language, left.sheet, row, column1);
                                if cell_status == CellTableStatus::Filtered {
                                    continue;
                                }
                                if mode == SubTotalMode::SkipHidden
                                    && cell_status == CellTableStatus::Hidden
                                {
                                    continue;
                                }
                                for column in column1..=column2 {
                                    if Model::cell_is_subtotal(workbook, cells, parsed_formulas, language, left.sheet, row, column) {
                                        continue;
                                    }
                                    match Model::evaluate_cell(workbook, cells, parsed_formulas, language, CellReferenceIndex {
                                        sheet: left.sheet,
                                        row,
                                        column,
                                    }) {
                                        CalcResult::EmptyCell | CalcResult::EmptyArg => {
                                            // skip
                                        }
                                        _ => counta += 1,
                                    }
                                }
                            }
                        }
                        CalcResult::String(_)
                        | CalcResult::Number(_)
                        | CalcResult::Boolean(_)
                        | CalcResult::Error { .. } => counta += 1,
                    }
                }
            }
        }
        CalcResult::Number(counta as f64)
    }

    fn subtotal_count(
        workbook: &Workbook,
        cells: &mut HashMap<(u32, i32, i32), CellState>,
        parsed_formulas: &Vec<Vec<Node>>,
        language: &Language,
        args: &[Node],
        cell: CellReferenceIndex,
        mode: SubTotalMode,
    ) -> CalcResult {
        let mut count = 0;
        for arg in args {
            match arg {
                Node::FunctionKind {
                    kind: Function::Subtotal,
                    args: _,
                } => {
                    // skip
                }
                _ => {
                    match Model::evaluate_node_with_reference(workbook, cells, parsed_formulas, language, arg, cell) {
                        CalcResult::Range { left, right } => {
                            if left.sheet != right.sheet {
                                return CalcResult::new_error(
                                    Error::VALUE,
                                    cell,
                                    "Ranges are in different sheets".to_string(),
                                );
                            }
                            // We are not expecting subtotal to have open ranges
                            let row1 = left.row;
                            let row2 = right.row;
                            let column1 = left.column;
                            let column2 = right.column;

                            for row in row1..=row2 {
                                let cell_status = Model::cell_hidden_status(workbook, cells, parsed_formulas, language, left.sheet, row, column1);
                                if cell_status == CellTableStatus::Filtered {
                                    continue;
                                }
                                if mode == SubTotalMode::SkipHidden
                                    && cell_status == CellTableStatus::Hidden
                                {
                                    continue;
                                }
                                for column in column1..=column2 {
                                    if Model::cell_is_subtotal(workbook, cells, parsed_formulas, language, left.sheet, row, column) {
                                        continue;
                                    }
                                    if let CalcResult::Number(_) =
                                        Model::evaluate_cell(workbook, cells, parsed_formulas, language, CellReferenceIndex {
                                            sheet: left.sheet,
                                            row,
                                            column,
                                        })
                                    {
                                        count += 1;
                                    }
                                }
                            }
                        }
                        // This hasn't been tested
                        CalcResult::Number(_) => count += 1,
                        _ => {}
                    }
                }
            }
        }
        CalcResult::Number(count as f64)
    }

    fn subtotal_average(
        workbook: &Workbook,
        cells: &mut HashMap<(u32, i32, i32), CellState>,
        parsed_formulas: &Vec<Vec<Node>>,
        language: &Language,
        args: &[Node],
        cell: CellReferenceIndex,
        mode: SubTotalMode,
    ) -> CalcResult {
        let values = match Model::subtotal_get_values(workbook, cells, parsed_formulas, language, args, cell, mode) {
            Ok(s) => s,
            Err(s) => return s,
        };
        let mut result = 0.0;
        let l = values.len();
        for value in values {
            result += value;
        }
        if l == 0 {
            return CalcResult::Error {
                error: Error::DIV,
                origin: cell,
                message: "Division by 0!".to_string(),
            };
        }
        CalcResult::Number(result / (l as f64))
    }

    fn subtotal_sum(
        workbook: &Workbook,
        cells: &mut HashMap<(u32, i32, i32), CellState>,
        parsed_formulas: &Vec<Vec<Node>>,
        language: &Language,
        args: &[Node],
        cell: CellReferenceIndex,
        mode: SubTotalMode,
    ) -> CalcResult {
        let values = match Model::subtotal_get_values(workbook, cells, parsed_formulas, language, args, cell, mode) {
            Ok(s) => s,
            Err(s) => return s,
        };
        let mut result = 0.0;
        for value in values {
            result += value;
        }
        CalcResult::Number(result)
    }

    fn subtotal_product(
        workbook: &Workbook,
        cells: &mut HashMap<(u32, i32, i32), CellState>,
        parsed_formulas: &Vec<Vec<Node>>,
        language: &Language,
        args: &[Node],
        cell: CellReferenceIndex,
        mode: SubTotalMode,
    ) -> CalcResult {
        let values = match Model::subtotal_get_values(workbook, cells, parsed_formulas, language, args, cell, mode) {
            Ok(s) => s,
            Err(s) => return s,
        };
        let mut result = 1.0;
        for value in values {
            result *= value;
        }
        CalcResult::Number(result)
    }

    fn subtotal_max(
        workbook: &Workbook,
        cells: &mut HashMap<(u32, i32, i32), CellState>,
        parsed_formulas: &Vec<Vec<Node>>,
        language: &Language,
        args: &[Node],
        cell: CellReferenceIndex,
        mode: SubTotalMode,
    ) -> CalcResult {
        let values = match Model::subtotal_get_values(workbook, cells, parsed_formulas, language, args, cell, mode) {
            Ok(s) => s,
            Err(s) => return s,
        };
        let mut result = f64::NAN;
        for value in values {
            result = value.max(result);
        }
        if result.is_nan() {
            return CalcResult::Number(0.0);
        }
        CalcResult::Number(result)
    }

    fn subtotal_min(
        workbook: &Workbook,
        cells: &mut HashMap<(u32, i32, i32), CellState>,
        parsed_formulas: &Vec<Vec<Node>>,
        language: &Language,
        args: &[Node],
        cell: CellReferenceIndex,
        mode: SubTotalMode,
    ) -> CalcResult {
        let values = match Model::subtotal_get_values(workbook, cells, parsed_formulas, language, args, cell, mode) {
            Ok(s) => s,
            Err(s) => return s,
        };
        let mut result = f64::NAN;
        for value in values {
            result = value.min(result);
        }
        if result.is_nan() {
            return CalcResult::Number(0.0);
        }
        CalcResult::Number(result)
    }
}
