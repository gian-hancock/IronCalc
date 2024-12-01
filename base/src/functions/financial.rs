use std::collections::HashMap;

use chrono::Datelike;

use crate::{
    calc_result::CalcResult, constants::{LAST_COLUMN, LAST_ROW}, expressions::{parser::Node, token::Error, types::CellReferenceIndex}, formatter::dates::from_excel_date, language::Language, model::Model, types::Workbook
};

use super::{financial_util::{compute_irr, compute_npv, compute_rate, compute_xirr, compute_xnpv}, CellState};

// See:
// https://github.com/apache/openoffice/blob/c014b5f2b55cff8d4b0c952d5c16d62ecde09ca1/main/scaddins/source/analysis/financial.cxx

// FIXME: Is this enough?
fn is_valid_date(date: f64) -> bool {
    date > 0.0
}

fn is_less_than_one_year(start_date: i64, end_date: i64) -> bool {
    if end_date - start_date < 365 {
        return true;
    }
    let end = from_excel_date(end_date);
    let start = from_excel_date(start_date);
    let end_year = end.year();
    let start_year = start.year();
    if end_year == start_year {
        return true;
    }
    if end_year != start_year + 1 {
        return false;
    }
    let start_month = start.month();
    let end_month = end.month();
    if end_month < start_month {
        return true;
    }
    if end_month > start_month {
        return false;
    }
    // we are one year later same month
    let start_day = start.day();
    let end_day = end.day();
    end_day <= start_day
}

fn compute_payment(
    rate: f64,
    nper: f64,
    pv: f64,
    fv: f64,
    period_start: bool,
) -> Result<f64, (Error, String)> {
    if rate == 0.0 {
        if nper == 0.0 {
            return Err((Error::NUM, "Period count must be non zero".to_string()));
        }
        return Ok(-(pv + fv) / nper);
    }
    if rate <= -1.0 {
        return Err((Error::NUM, "Rate must be > -1".to_string()));
    };
    let rate_nper = if nper == 0.0 {
        1.0
    } else {
        (1.0 + rate).powf(nper)
    };
    let result = if period_start {
        // type = 1
        (fv + pv * rate_nper) * rate / ((1.0 + rate) * (1.0 - rate_nper))
    } else {
        (fv * rate + pv * rate * rate_nper) / (1.0 - rate_nper)
    };
    if result.is_nan() || result.is_infinite() {
        return Err((Error::NUM, "Invalid result".to_string()));
    }
    Ok(result)
}

fn compute_future_value(
    rate: f64,
    nper: f64,
    pmt: f64,
    pv: f64,
    period_start: bool,
) -> Result<f64, (Error, String)> {
    if rate == 0.0 {
        return Ok(-pv - pmt * nper);
    }

    let rate_nper = (1.0 + rate).powf(nper);
    let fv = if period_start {
        // type = 1
        -pv * rate_nper - pmt * (1.0 + rate) * (rate_nper - 1.0) / rate
    } else {
        -pv * rate_nper - pmt * (rate_nper - 1.0) / rate
    };
    if fv.is_nan() {
        return Err((Error::NUM, "Invalid result".to_string()));
    }
    if !fv.is_finite() {
        return Err((Error::DIV, "Divide by zero".to_string()));
    }
    Ok(fv)
}

fn compute_ipmt(
    rate: f64,
    period: f64,
    period_count: f64,
    present_value: f64,
    future_value: f64,
    period_start: bool,
) -> Result<f64, (Error, String)> {
    // http://www.staff.city.ac.uk/o.s.kerr/CompMaths/WSheet4.pdf
    // https://www.experts-exchange.com/articles/1948/A-Guide-to-the-PMT-FV-IPMT-and-PPMT-Functions.html
    // type = 0 (end of period)
    // impt = -[(1+rate)^(period-1)*(pv*rate+pmt)-pmt]
    // ipmt = FV(rate, period-1, payment, pv, type) * rate
    // type = 1 (beginning of period)
    // ipmt = (FV(rate, period-2, payment, pv, type) - payment) * rate
    let payment = compute_payment(
        rate,
        period_count,
        present_value,
        future_value,
        period_start,
    )?;
    if period < 1.0 || period >= period_count + 1.0 {
        return Err((
            Error::NUM,
            format!("Period must be between 1 and {}", period_count + 1.0),
        ));
    }
    if period == 1.0 && period_start {
        Ok(0.0)
    } else {
        let p = if period_start {
            period - 2.0
        } else {
            period - 1.0
        };
        let c = if period_start { -payment } else { 0.0 };
        let fv = compute_future_value(rate, p, payment, present_value, period_start)?;
        Ok((fv + c) * rate)
    }
}

fn compute_ppmt(
    rate: f64,
    period: f64,
    period_count: f64,
    present_value: f64,
    future_value: f64,
    period_start: bool,
) -> Result<f64, (Error, String)> {
    let payment = compute_payment(
        rate,
        period_count,
        present_value,
        future_value,
        period_start,
    )?;
    // It's a bit unfortunate that the first thing compute_ipmt does is compute_payment again
    let ipmt = compute_ipmt(
        rate,
        period,
        period_count,
        present_value,
        future_value,
        period_start,
    )?;
    Ok(payment - ipmt)
}

// These formulas revolve around compound interest and annuities.
// The financial functions pv, rate, nper, pmt and fv:
// rate = interest rate per period
// nper (number of periods) = loan term
// pv (present value) = loan amount
// fv (future value) = cash balance after last payment. Default is 0
// type = the annuity type indicates when payments are due
//         * 0 (default) Payments are made at the end of the period
//         * 1 Payments are made at the beginning of the period (like a lease or rent)
// The variable period_start is true if type is 1
// They are linked by the formulas:
// If rate != 0
//   $pv*(1+rate)^nper+pmt*(1+rate*type)*((1+rate)^nper-1)/rate+fv=0$
// If rate = 0
//   $pmt*nper+pv+fv=0$
// All, except for rate are easily solvable in terms of the others.
// In these formulas the payment (pmt) is normally negative

impl Model {
    // FIXME: These three functions (get_array_of_numbers..) need to be refactored
    // They are really similar expect for small issues
    fn get_array_of_numbers(
        workbook: &Workbook,
        cells: &mut HashMap<(u32, i32, i32), CellState>,
        parsed_formulas: &Vec<Vec<Node>>,
        language: &Language,
        arg: &Node,
        cell: &CellReferenceIndex,
    ) -> Result<Vec<f64>, CalcResult> {
        let mut values = Vec::new();
        match Model::evaluate_node_in_context(workbook, cells, parsed_formulas, language, arg, *cell) {
            CalcResult::Number(value) => values.push(value),
            CalcResult::Range { left, right } => {
                if left.sheet != right.sheet {
                    return Err(CalcResult::new_error(
                        Error::VALUE,
                        *cell,
                        "Ranges are in different sheets".to_string(),
                    ));
                }
                let row1 = left.row;
                let mut row2 = right.row;
                let column1 = left.column;
                let mut column2 = right.column;
                if row1 == 1 && row2 == LAST_ROW {
                    row2 = workbook
                        .worksheet(left.sheet)
                        .expect("Sheet expected during evaluation.")
                        .dimension()
                        .max_row;
                }
                if column1 == 1 && column2 == LAST_COLUMN {
                    column2 = workbook
                        .worksheet(left.sheet)
                        .expect("Sheet expected during evaluation.")
                        .dimension()
                        .max_column;
                }
                for row in row1..row2 + 1 {
                    for column in column1..(column2 + 1) {
                        match Model::evaluate_cell(workbook, cells, parsed_formulas, language, CellReferenceIndex {
                            sheet: left.sheet,
                            row,
                            column,
                        }) {
                            CalcResult::Number(value) => {
                                values.push(value);
                            }
                            error @ CalcResult::Error { .. } => return Err(error),
                            _ => {
                                // We ignore booleans and strings
                            }
                        }
                    }
                }
            }
            error @ CalcResult::Error { .. } => return Err(error),
            _ => {
                // We ignore booleans and strings
            }
        };
        Ok(values)
    }

    fn get_array_of_numbers_xpnv(
        workbook: &Workbook,
        cells: &mut HashMap<(u32, i32, i32), CellState>,
        parsed_formulas: &Vec<Vec<Node>>,
        language: &Language,
        cell: &CellReferenceIndex,
        arg: &Node,
        error: Error,
    ) -> Result<Vec<f64>, CalcResult> {
        let mut values = Vec::new();
        match Model::evaluate_node_in_context(workbook, cells, parsed_formulas, language, arg, *cell) {
            CalcResult::Number(value) => values.push(value),
            CalcResult::Range { left, right } => {
                if left.sheet != right.sheet {
                    return Err(CalcResult::new_error(
                        Error::VALUE,
                        *cell,
                        "Ranges are in different sheets".to_string(),
                    ));
                }
                let row1 = left.row;
                let mut row2 = right.row;
                let column1 = left.column;
                let mut column2 = right.column;
                if row1 == 1 && row2 == LAST_ROW {
                    row2 = workbook
                        .worksheet(left.sheet)
                        .expect("Sheet expected during evaluation.")
                        .dimension()
                        .max_row;
                }
                if column1 == 1 && column2 == LAST_COLUMN {
                    column2 = workbook
                        .worksheet(left.sheet)
                        .expect("Sheet expected during evaluation.")
                        .dimension()
                        .max_column;
                }
                for row in row1..row2 + 1 {
                    for column in column1..(column2 + 1) {
                        match Model::evaluate_cell(workbook, cells, parsed_formulas, language, CellReferenceIndex {
                            sheet: left.sheet,
                            row,
                            column,
                        }) {
                            CalcResult::Number(value) => {
                                values.push(value);
                            }
                            error @ CalcResult::Error { .. } => return Err(error),
                            CalcResult::EmptyCell => {
                                return Err(CalcResult::new_error(
                                    Error::NUM,
                                    *cell,
                                    "Expected number".to_string(),
                                ));
                            }
                            _ => {
                                return Err(CalcResult::new_error(
                                    error,
                                    *cell,
                                    "Expected number".to_string(),
                                ));
                            }
                        }
                    }
                }
            }
            error @ CalcResult::Error { .. } => return Err(error),
            _ => {
                return Err(CalcResult::new_error(
                    error,
                    *cell,
                    "Expected number".to_string(),
                ));
            }
        };
        Ok(values)
    }

    fn get_array_of_numbers_xirr(
        workbook: &Workbook,
        cells: &mut HashMap<(u32, i32, i32), CellState>,
        parsed_formulas: &Vec<Vec<Node>>,
        language: &Language,
        arg: &Node,
        cell: &CellReferenceIndex,
    ) -> Result<Vec<f64>, CalcResult> {
        let mut values = Vec::new();
        match Model::evaluate_node_in_context(workbook, cells, parsed_formulas, language, arg, *cell) {
            CalcResult::Range { left, right } => {
                if left.sheet != right.sheet {
                    return Err(CalcResult::new_error(
                        Error::VALUE,
                        *cell,
                        "Ranges are in different sheets".to_string(),
                    ));
                }
                let row1 = left.row;
                let mut row2 = right.row;
                let column1 = left.column;
                let mut column2 = right.column;
                if row1 == 1 && row2 == LAST_ROW {
                    row2 = workbook
                        .worksheet(left.sheet)
                        .expect("Sheet expected during evaluation.")
                        .dimension()
                        .max_row;
                }
                if column1 == 1 && column2 == LAST_COLUMN {
                    column2 = workbook
                        .worksheet(left.sheet)
                        .expect("Sheet expected during evaluation.")
                        .dimension()
                        .max_column;
                }
                for row in row1..row2 + 1 {
                    for column in column1..(column2 + 1) {
                        match Model::evaluate_cell(workbook, cells, parsed_formulas, language, CellReferenceIndex {
                            sheet: left.sheet,
                            row,
                            column,
                        }) {
                            CalcResult::Number(value) => {
                                values.push(value);
                            }
                            error @ CalcResult::Error { .. } => return Err(error),
                            CalcResult::EmptyCell => values.push(0.0),
                            _ => {
                                return Err(CalcResult::new_error(
                                    Error::VALUE,
                                    *cell,
                                    "Expected number".to_string(),
                                ));
                            }
                        }
                    }
                }
            }
            error @ CalcResult::Error { .. } => return Err(error),
            _ => {
                return Err(CalcResult::new_error(
                    Error::VALUE,
                    *cell,
                    "Expected number".to_string(),
                ));
            }
        };
        Ok(values)
    }

    /// PMT(rate, nper, pv, [fv], [type])
    pub(crate) fn fn_pmt(
        workbook: &Workbook,
        cells: &mut HashMap<(u32, i32, i32), CellState>,
        parsed_formulas: &Vec<Vec<Node>>,
        language: &Language,
        args: &[Node], cell: CellReferenceIndex) -> CalcResult {
        let arg_count = args.len();
        if !(3..=5).contains(&arg_count) {
            return CalcResult::new_args_number_error(cell);
        }
        let rate = match Model::get_number(workbook, cells, parsed_formulas, language, &args[0], cell) {
            Ok(f) => f,
            Err(s) => return s,
        };
        // number of periods
        let nper = match Model::get_number(workbook, cells, parsed_formulas, language, &args[1], cell) {
            Ok(f) => f,
            Err(s) => return s,
        };
        // present value
        let pv = match Model::get_number(workbook, cells, parsed_formulas, language, &args[2], cell) {
            Ok(f) => f,
            Err(s) => return s,
        };
        // future_value
        let fv = if arg_count > 3 {
            match Model::get_number(workbook, cells, parsed_formulas, language, &args[3], cell) {
                Ok(f) => f,
                Err(s) => return s,
            }
        } else {
            0.0
        };
        let period_start = if arg_count > 4 {
            match Model::get_number(workbook, cells, parsed_formulas, language, &args[4], cell) {
                Ok(f) => f != 0.0,
                Err(s) => return s,
            }
        } else {
            // at the end of the period
            false
        };
        match compute_payment(rate, nper, pv, fv, period_start) {
            Ok(p) => CalcResult::Number(p),
            Err(error) => CalcResult::Error {
                error: error.0,
                origin: cell,
                message: error.1,
            },
        }
    }

    // PV(rate, nper, pmt, [fv], [type])
    pub(crate) fn fn_pv(
        workbook: &Workbook,
        cells: &mut HashMap<(u32, i32, i32), CellState>,
        parsed_formulas: &Vec<Vec<Node>>,
        language: &Language,
        args: &[Node], cell: CellReferenceIndex) -> CalcResult {
        let arg_count = args.len();
        if !(3..=5).contains(&arg_count) {
            return CalcResult::new_args_number_error(cell);
        }
        let rate = match Model::get_number(workbook, cells, parsed_formulas, language, &args[0], cell) {
            Ok(f) => f,
            Err(s) => return s,
        };
        // nper
        let period_count = match Model::get_number(workbook, cells, parsed_formulas, language, &args[1], cell) {
            Ok(f) => f,
            Err(s) => return s,
        };
        // pmt
        let payment = match Model::get_number(workbook, cells, parsed_formulas, language, &args[2], cell) {
            Ok(f) => f,
            Err(s) => return s,
        };
        // fv
        let future_value = if arg_count > 3 {
            match Model::get_number(workbook, cells, parsed_formulas, language, &args[3], cell) {
                Ok(f) => f,
                Err(s) => return s,
            }
        } else {
            0.0
        };
        let period_start = if arg_count > 4 {
            match Model::get_number(workbook, cells, parsed_formulas, language, &args[4], cell) {
                Ok(f) => f != 0.0,
                Err(s) => return s,
            }
        } else {
            // at the end of the period
            false
        };
        if rate == 0.0 {
            return CalcResult::Number(-future_value - payment * period_count);
        }
        if rate == -1.0 {
            return CalcResult::Error {
                error: Error::NUM,
                origin: cell,
                message: "Rate must be != -1".to_string(),
            };
        };
        let rate_nper = (1.0 + rate).powf(period_count);
        let result = if period_start {
            // type = 1
            -(future_value * rate + payment * (1.0 + rate) * (rate_nper - 1.0)) / (rate * rate_nper)
        } else {
            (-future_value * rate - payment * (rate_nper - 1.0)) / (rate * rate_nper)
        };
        if result.is_nan() || result.is_infinite() {
            return CalcResult::Error {
                error: Error::NUM,
                origin: cell,
                message: "Invalid result".to_string(),
            };
        }

        CalcResult::Number(result)
    }

    // RATE(nper, pmt, pv, [fv], [type], [guess])
    pub(crate) fn fn_rate(
        workbook: &Workbook,
        cells: &mut HashMap<(u32, i32, i32), CellState>,
        parsed_formulas: &Vec<Vec<Node>>,
        language: &Language,
        args: &[Node], cell: CellReferenceIndex) -> CalcResult {
        let arg_count = args.len();
        if !(3..=5).contains(&arg_count) {
            return CalcResult::new_args_number_error(cell);
        }
        let nper = match Model::get_number(workbook, cells, parsed_formulas, language, &args[0], cell) {
            Ok(f) => f,
            Err(s) => return s,
        };
        let pmt = match Model::get_number(workbook, cells, parsed_formulas, language, &args[1], cell) {
            Ok(f) => f,
            Err(s) => return s,
        };
        let pv = match Model::get_number(workbook, cells, parsed_formulas, language, &args[2], cell) {
            Ok(f) => f,
            Err(s) => return s,
        };
        // fv
        let fv = if arg_count > 3 {
            match Model::get_number(workbook, cells, parsed_formulas, language, &args[3], cell) {
                Ok(f) => f,
                Err(s) => return s,
            }
        } else {
            0.0
        };
        let annuity_type = if arg_count > 4 {
            match Model::get_number(workbook, cells, parsed_formulas, language, &args[4], cell) {
                Ok(f) => i32::from(f != 0.0),
                Err(s) => return s,
            }
        } else {
            // at the end of the period
            0
        };

        let guess = if arg_count > 5 {
            match Model::get_number(workbook, cells, parsed_formulas, language, &args[5], cell) {
                Ok(f) => f,
                Err(s) => return s,
            }
        } else {
            0.1
        };

        match compute_rate(pv, fv, nper, pmt, annuity_type, guess) {
            Ok(f) => CalcResult::Number(f),
            Err(error) => CalcResult::Error {
                error: error.0,
                origin: cell,
                message: error.1,
            },
        }
    }

    // NPER(rate,pmt,pv,[fv],[type])
    pub(crate) fn fn_nper(
        workbook: &Workbook,
        cells: &mut HashMap<(u32, i32, i32), CellState>,
        parsed_formulas: &Vec<Vec<Node>>,
        language: &Language,
        args: &[Node], cell: CellReferenceIndex) -> CalcResult {
        let arg_count = args.len();
        if !(3..=5).contains(&arg_count) {
            return CalcResult::new_args_number_error(cell);
        }
        let rate = match Model::get_number(workbook, cells, parsed_formulas, language, &args[0], cell) {
            Ok(f) => f,
            Err(s) => return s,
        };
        // pmt
        let payment = match Model::get_number(workbook, cells, parsed_formulas, language, &args[1], cell) {
            Ok(f) => f,
            Err(s) => return s,
        };
        // pv
        let present_value = match Model::get_number(workbook, cells, parsed_formulas, language, &args[2], cell) {
            Ok(f) => f,
            Err(s) => return s,
        };
        // fv
        let future_value = if arg_count > 3 {
            match Model::get_number(workbook, cells, parsed_formulas, language, &args[3], cell) {
                Ok(f) => f,
                Err(s) => return s,
            }
        } else {
            0.0
        };
        let period_start = if arg_count > 4 {
            match Model::get_number(workbook, cells, parsed_formulas, language, &args[4], cell) {
                Ok(f) => f != 0.0,
                Err(s) => return s,
            }
        } else {
            // at the end of the period
            false
        };
        if rate == 0.0 {
            if payment == 0.0 {
                return CalcResult::Error {
                    error: Error::DIV,
                    origin: cell,
                    message: "Divide by zero".to_string(),
                };
            }
            return CalcResult::Number(-(future_value + present_value) / payment);
        }
        if rate < -1.0 {
            return CalcResult::Error {
                error: Error::NUM,
                origin: cell,
                message: "Rate must be > -1".to_string(),
            };
        };
        let rate_nper = if period_start {
            // type = 1
            if payment != 0.0 {
                let term = payment * (1.0 + rate) / rate;
                (1.0 - future_value / term) / (1.0 + present_value / term)
            } else {
                -future_value / present_value
            }
        } else {
            // type = 0
            if payment != 0.0 {
                let term = payment / rate;
                (1.0 - future_value / term) / (1.0 + present_value / term)
            } else {
                -future_value / present_value
            }
        };
        if rate_nper <= 0.0 {
            return CalcResult::Error {
                error: Error::NUM,
                origin: cell,
                message: "Cannot compute.".to_string(),
            };
        }
        let result = rate_nper.ln() / (1.0 + rate).ln();
        CalcResult::Number(result)
    }

    // FV(rate, nper, pmt, [pv], [type])
    pub(crate) fn fn_fv(
        workbook: &Workbook,
        cells: &mut HashMap<(u32, i32, i32), CellState>,
        parsed_formulas: &Vec<Vec<Node>>,
        language: &Language,
        args: &[Node], cell: CellReferenceIndex) -> CalcResult {
        let arg_count = args.len();
        if !(3..=5).contains(&arg_count) {
            return CalcResult::new_args_number_error(cell);
        }
        let rate = match Model::get_number(workbook, cells, parsed_formulas, language, &args[0], cell) {
            Ok(f) => f,
            Err(s) => return s,
        };
        // number of periods
        let nper = match Model::get_number(workbook, cells, parsed_formulas, language, &args[1], cell) {
            Ok(f) => f,
            Err(s) => return s,
        };
        // payment
        let pmt = match Model::get_number(workbook, cells, parsed_formulas, language, &args[2], cell) {
            Ok(f) => f,
            Err(s) => return s,
        };
        // present value
        let pv = if arg_count > 3 {
            match Model::get_number(workbook, cells, parsed_formulas, language, &args[3], cell) {
                Ok(f) => f,
                Err(s) => return s,
            }
        } else {
            0.0
        };
        let period_start = if arg_count > 4 {
            match Model::get_number(workbook, cells, parsed_formulas, language, &args[4], cell) {
                Ok(f) => f != 0.0,
                Err(s) => return s,
            }
        } else {
            // at the end of the period
            false
        };
        match compute_future_value(rate, nper, pmt, pv, period_start) {
            Ok(f) => CalcResult::Number(f),
            Err(error) => CalcResult::Error {
                error: error.0,
                origin: cell,
                message: error.1,
            },
        }
    }

    // IPMT(rate, per, nper, pv, [fv], [type])
    pub(crate) fn fn_ipmt(
        workbook: &Workbook,
        cells: &mut HashMap<(u32, i32, i32), CellState>,
        parsed_formulas: &Vec<Vec<Node>>,
        language: &Language,
        args: &[Node], cell: CellReferenceIndex) -> CalcResult {
        let arg_count = args.len();
        if !(4..=6).contains(&arg_count) {
            return CalcResult::new_args_number_error(cell);
        }
        let rate = match Model::get_number(workbook, cells, parsed_formulas, language, &args[0], cell) {
            Ok(f) => f,
            Err(s) => return s,
        };
        // per
        let period = match Model::get_number(workbook, cells, parsed_formulas, language, &args[1], cell) {
            Ok(f) => f,
            Err(s) => return s,
        };
        // nper
        let period_count = match Model::get_number(workbook, cells, parsed_formulas, language, &args[2], cell) {
            Ok(f) => f,
            Err(s) => return s,
        };
        // pv
        let present_value = match Model::get_number(workbook, cells, parsed_formulas, language, &args[3], cell) {
            Ok(f) => f,
            Err(s) => return s,
        };
        // fv
        let future_value = if arg_count > 4 {
            match Model::get_number(workbook, cells, parsed_formulas, language, &args[4], cell) {
                Ok(f) => f,
                Err(s) => return s,
            }
        } else {
            0.0
        };
        let period_start = if arg_count > 5 {
            match Model::get_number(workbook, cells, parsed_formulas, language, &args[5], cell) {
                Ok(f) => f != 0.0,
                Err(s) => return s,
            }
        } else {
            // at the end of the period
            false
        };
        let ipmt = match compute_ipmt(
            rate,
            period,
            period_count,
            present_value,
            future_value,
            period_start,
        ) {
            Ok(f) => f,
            Err(error) => {
                return CalcResult::Error {
                    error: error.0,
                    origin: cell,
                    message: error.1,
                }
            }
        };
        CalcResult::Number(ipmt)
    }

    // PPMT(rate, per, nper, pv, [fv], [type])
    pub(crate) fn fn_ppmt(
        workbook: &Workbook,
        cells: &mut HashMap<(u32, i32, i32), CellState>,
        parsed_formulas: &Vec<Vec<Node>>,
        language: &Language,
        args: &[Node], cell: CellReferenceIndex) -> CalcResult {
        let arg_count = args.len();
        if !(4..=6).contains(&arg_count) {
            return CalcResult::new_args_number_error(cell);
        }
        let rate = match Model::get_number(workbook, cells, parsed_formulas, language, &args[0], cell) {
            Ok(f) => f,
            Err(s) => return s,
        };
        // per
        let period = match Model::get_number(workbook, cells, parsed_formulas, language, &args[1], cell) {
            Ok(f) => f,
            Err(s) => return s,
        };
        // nper
        let period_count = match Model::get_number(workbook, cells, parsed_formulas, language, &args[2], cell) {
            Ok(f) => f,
            Err(s) => return s,
        };
        // pv
        let present_value = match Model::get_number(workbook, cells, parsed_formulas, language, &args[3], cell) {
            Ok(f) => f,
            Err(s) => return s,
        };
        // fv
        let future_value = if arg_count > 4 {
            match Model::get_number(workbook, cells, parsed_formulas, language, &args[4], cell) {
                Ok(f) => f,
                Err(s) => return s,
            }
        } else {
            0.0
        };
        let period_start = if arg_count > 5 {
            match Model::get_number(workbook, cells, parsed_formulas, language, &args[5], cell) {
                Ok(f) => f != 0.0,
                Err(s) => return s,
            }
        } else {
            // at the end of the period
            false
        };

        let ppmt = match compute_ppmt(
            rate,
            period,
            period_count,
            present_value,
            future_value,
            period_start,
        ) {
            Ok(f) => f,
            Err(error) => {
                return CalcResult::Error {
                    error: error.0,
                    origin: cell,
                    message: error.1,
                }
            }
        };
        CalcResult::Number(ppmt)
    }

    // NPV(rate, value1, [value2],...)
    // npv = Sum[value[i]/(1+rate)^i, {i, 1, n}]
    pub(crate) fn fn_npv(
        workbook: &Workbook,
        cells: &mut HashMap<(u32, i32, i32), CellState>,
        parsed_formulas: &Vec<Vec<Node>>,
        language: &Language,
        args: &[Node], cell: CellReferenceIndex) -> CalcResult {
        let arg_count = args.len();
        if arg_count < 2 {
            return CalcResult::new_args_number_error(cell);
        }
        let rate = match Model::get_number(workbook, cells, parsed_formulas, language, &args[0], cell) {
            Ok(f) => f,
            Err(s) => return s,
        };
        let mut values = Vec::new();
        for arg in &args[1..] {
            match Model::evaluate_node_in_context(workbook, cells, parsed_formulas, language, arg, cell) {
                CalcResult::Number(value) => values.push(value),
                CalcResult::Range { left, right } => {
                    if left.sheet != right.sheet {
                        return CalcResult::new_error(
                            Error::VALUE,
                            cell,
                            "Ranges are in different sheets".to_string(),
                        );
                    }
                    let row1 = left.row;
                    let mut row2 = right.row;
                    let column1 = left.column;
                    let mut column2 = right.column;
                    if row1 == 1 && row2 == LAST_ROW {
                        row2 = workbook
                            .worksheet(left.sheet)
                            .expect("Sheet expected during evaluation.")
                            .dimension()
                            .max_row;
                    }
                    if column1 == 1 && column2 == LAST_COLUMN {
                        column2 = workbook
                            .worksheet(left.sheet)
                            .expect("Sheet expected during evaluation.")
                            .dimension()
                            .max_column;
                    }
                    for row in row1..row2 + 1 {
                        for column in column1..(column2 + 1) {
                            match Model::evaluate_cell(workbook, cells, parsed_formulas, language, CellReferenceIndex {
                                sheet: left.sheet,
                                row,
                                column,
                            }) {
                                CalcResult::Number(value) => {
                                    values.push(value);
                                }
                                error @ CalcResult::Error { .. } => return error,
                                _ => {
                                    // We ignore booleans and strings
                                }
                            }
                        }
                    }
                }
                error @ CalcResult::Error { .. } => return error,
                _ => {
                    // We ignore booleans and strings
                }
            };
        }
        match compute_npv(rate, &values) {
            Ok(f) => CalcResult::Number(f),
            Err(error) => CalcResult::new_error(error.0, cell, error.1),
        }
    }

    // Returns the internal rate of return for a series of cash flows represented by the numbers
    // in values.
    // These cash flows do not have to be even, as they would be for an annuity.
    // However, the cash flows must occur at regular intervals, such as monthly or annually.
    // The internal rate of return is the interest rate received for an investment consisting
    // of payments (negative values) and income (positive values) that occur at regular periods

    // IRR(values, [guess])
    pub(crate) fn fn_irr(
        workbook: &Workbook,
        cells: &mut HashMap<(u32, i32, i32), CellState>,
        parsed_formulas: &Vec<Vec<Node>>,
        language: &Language,
        args: &[Node], cell: CellReferenceIndex) -> CalcResult {
        let arg_count = args.len();
        if arg_count > 2 || arg_count == 0 {
            return CalcResult::new_args_number_error(cell);
        }
        let values = match Model::get_array_of_numbers(workbook, cells, parsed_formulas, language, &args[0], &cell) {
            Ok(s) => s,
            Err(error) => return error,
        };
        let guess = if arg_count == 2 {
            match Model::get_number(workbook, cells, parsed_formulas, language, &args[1], cell) {
                Ok(f) => f,
                Err(s) => return s,
            }
        } else {
            0.1
        };
        match compute_irr(&values, guess) {
            Ok(f) => CalcResult::Number(f),
            Err(error) => CalcResult::Error {
                error: error.0,
                origin: cell,
                message: error.1,
            },
        }
    }

    // XNPV(rate, values, dates)
    pub(crate) fn fn_xnpv(
        workbook: &Workbook,
        cells: &mut HashMap<(u32, i32, i32), CellState>,
        parsed_formulas: &Vec<Vec<Node>>,
        language: &Language,
        args: &[Node], cell: CellReferenceIndex) -> CalcResult {
        let arg_count = args.len();
        if !(2..=3).contains(&arg_count) {
            return CalcResult::new_args_number_error(cell);
        }
        let rate = match Model::get_number(workbook, cells, parsed_formulas, language, &args[0], cell) {
            Ok(f) => f,
            Err(s) => return s,
        };
        let values = match Model::get_array_of_numbers_xpnv(workbook, cells,parsed_formulas, language, &cell, &args[1], Error::NUM) {
            Ok(s) => s,
            Err(error) => return error,
        };
        let dates = match Model::get_array_of_numbers_xpnv(workbook, cells, parsed_formulas, language, &cell, &args[2], Error::VALUE) {
            Ok(s) => s,
            Err(error) => return error,
        };
        // Decimal points on dates are truncated
        let dates: Vec<f64> = dates.iter().map(|s| s.floor()).collect();
        let values_count = values.len();
        // If values and dates contain a different number of values, XNPV returns the #NUM! error value.
        if values_count != dates.len() {
            return CalcResult::new_error(
                Error::NUM,
                cell,
                "Values and dates must be the same length".to_string(),
            );
        }
        if values_count == 0 {
            return CalcResult::new_error(Error::NUM, cell, "Not enough values".to_string());
        }
        let first_date = dates[0];
        for date in &dates {
            if !is_valid_date(*date) {
                // Excel docs claim that if any number in dates is not a valid date,
                // XNPV returns the #VALUE! error value, but it seems to return #VALUE!
                return CalcResult::new_error(
                    Error::NUM,
                    cell,
                    "Invalid number for date".to_string(),
                );
            }
            // If any number in dates precedes the starting date, XNPV returns the #NUM! error value.
            if date < &first_date {
                return CalcResult::new_error(
                    Error::NUM,
                    cell,
                    "Date precedes the starting date".to_string(),
                );
            }
        }
        // It seems Excel returns #NUM! if rate < 0, this is only necessary if r <= -1
        if rate <= 0.0 {
            return CalcResult::new_error(Error::NUM, cell, "rate needs to be > 0".to_string());
        }
        match compute_xnpv(rate, &values, &dates) {
            Ok(f) => CalcResult::Number(f),
            Err((error, message)) => CalcResult::new_error(error, cell, message),
        }
    }

    // XIRR(values, dates, [guess])
    pub(crate) fn fn_xirr(
        workbook: &Workbook,
        cells: &mut HashMap<(u32, i32, i32), CellState>,
        parsed_formulas: &Vec<Vec<Node>>,
        language: &Language,
        args: &[Node], cell: CellReferenceIndex) -> CalcResult {
        let arg_count = args.len();
        if !(2..=3).contains(&arg_count) {
            return CalcResult::new_args_number_error(cell);
        }
        let values = match Model::get_array_of_numbers_xirr(workbook, cells, parsed_formulas, language, &args[0], &cell) {
            Ok(s) => s,
            Err(error) => return error,
        };
        let dates = match Model::get_array_of_numbers_xirr(workbook, cells, parsed_formulas, language, &args[1], &cell) {
            Ok(s) => s,
            Err(error) => return error,
        };
        let guess = if arg_count == 3 {
            match Model::get_number(workbook, cells, parsed_formulas, language, &args[2], cell) {
                Ok(f) => f,
                Err(s) => return s,
            }
        } else {
            0.1
        };
        // Decimal points on dates are truncated
        let dates: Vec<f64> = dates.iter().map(|s| s.floor()).collect();
        let values_count = values.len();
        // If values and dates contain a different number of values, XNPV returns the #NUM! error value.
        if values_count != dates.len() {
            return CalcResult::new_error(
                Error::NUM,
                cell,
                "Values and dates must be the same length".to_string(),
            );
        }
        if values_count == 0 {
            return CalcResult::new_error(Error::NUM, cell, "Not enough values".to_string());
        }
        let first_date = dates[0];
        for date in &dates {
            if !is_valid_date(*date) {
                return CalcResult::new_error(
                    Error::NUM,
                    cell,
                    "Invalid number for date".to_string(),
                );
            }
            // If any number in dates precedes the starting date, XIRR returns the #NUM! error value.
            if date < &first_date {
                return CalcResult::new_error(
                    Error::NUM,
                    cell,
                    "Date precedes the starting date".to_string(),
                );
            }
        }
        match compute_xirr(&values, &dates, guess) {
            Ok(f) => CalcResult::Number(f),
            Err((error, message)) => CalcResult::Error {
                error,
                origin: cell,
                message,
            },
        }
    }

    //  MIRR(values, finance_rate, reinvest_rate)
    // The formula is:
    // $$ (-NPV(r1, v_p) * (1+r1)^y)/(NPV(r2, v_n)*(1+r2))^(1/y)-1$$
    // where:
    // $r1$ is the reinvest_rate, $r2$ the finance_rate
    // $v_p$ the vector of positive values
    // $v_n$ the vector of negative values
    // and $y$ is dimension of $v$ - 1 (number of years)
    pub(crate) fn fn_mirr(
        workbook: &Workbook,
        cells: &mut HashMap<(u32, i32, i32), CellState>,
        parsed_formulas: &Vec<Vec<Node>>,
        language: &Language,
        args: &[Node], cell: CellReferenceIndex) -> CalcResult {
        if args.len() != 3 {
            return CalcResult::new_args_number_error(cell);
        }
        let values = match Model::get_array_of_numbers(workbook, cells, parsed_formulas, language, &args[0], &cell) {
            Ok(s) => s,
            Err(error) => return error,
        };
        let finance_rate = match Model::get_number(workbook, cells, parsed_formulas, language, &args[1], cell) {
            Ok(f) => f,
            Err(s) => return s,
        };
        let reinvest_rate = match Model::get_number(workbook, cells, parsed_formulas, language, &args[2], cell) {
            Ok(f) => f,
            Err(s) => return s,
        };
        let mut positive_values = Vec::new();
        let mut negative_values = Vec::new();
        let mut last_negative_index = -1;
        for (index, &value) in values.iter().enumerate() {
            let (p, n) = if value >= 0.0 {
                (value, 0.0)
            } else {
                last_negative_index = index as i32;
                (0.0, value)
            };
            positive_values.push(p);
            negative_values.push(n);
        }
        if last_negative_index == -1 {
            return CalcResult::new_error(
                Error::DIV,
                cell,
                "Invalid data for MIRR function".to_string(),
            );
        }
        // We do a bit of analysis if the rates are -1 as there are some cancellations
        // It is probably not important.
        let years = values.len() as f64;
        let top = if reinvest_rate == -1.0 {
            // This is finite
            match positive_values.last() {
                Some(f) => *f,
                None => 0.0,
            }
        } else {
            match compute_npv(reinvest_rate, &positive_values) {
                Ok(npv) => -npv * ((1.0 + reinvest_rate).powf(years)),
                Err((error, message)) => {
                    return CalcResult::Error {
                        error,
                        origin: cell,
                        message,
                    }
                }
            }
        };
        let bottom = if finance_rate == -1.0 {
            if last_negative_index == 0 {
                // This is still finite
                negative_values[last_negative_index as usize]
            } else {
                // or -Infinity depending of the sign in the last_negative_index coef.
                // But it is irrelevant for the calculation
                f64::INFINITY
            }
        } else {
            match compute_npv(finance_rate, &negative_values) {
                Ok(npv) => npv * (1.0 + finance_rate),
                Err((error, message)) => {
                    return CalcResult::Error {
                        error,
                        origin: cell,
                        message,
                    }
                }
            }
        };

        let result = (top / bottom).powf(1.0 / (years - 1.0)) - 1.0;
        if result.is_infinite() {
            return CalcResult::new_error(Error::DIV, cell, "Division by 0".to_string());
        }
        if result.is_nan() {
            return CalcResult::new_error(Error::NUM, cell, "Invalid data for MIRR".to_string());
        }
        CalcResult::Number(result)
    }

    // ISPMT(rate, per, nper, pv)
    // Formula is:
    // $$pv*rate*\left(\frac{per}{nper}-1\right)$$
    pub(crate) fn fn_ispmt(
        workbook: &Workbook,
        cells: &mut HashMap<(u32, i32, i32), CellState>,
        parsed_formulas: &Vec<Vec<Node>>,
        language: &Language,
        args: &[Node], cell: CellReferenceIndex) -> CalcResult {
        if args.len() != 4 {
            return CalcResult::new_args_number_error(cell);
        }
        let rate = match Model::get_number(workbook, cells, parsed_formulas, language, &args[0], cell) {
            Ok(f) => f,
            Err(s) => return s,
        };
        let per = match Model::get_number(workbook, cells, parsed_formulas, language, &args[1], cell) {
            Ok(f) => f,
            Err(s) => return s,
        };
        let nper = match Model::get_number(workbook, cells, parsed_formulas, language, &args[2], cell) {
            Ok(f) => f,
            Err(s) => return s,
        };
        let pv = match Model::get_number(workbook, cells, parsed_formulas, language, &args[3], cell) {
            Ok(f) => f,
            Err(s) => return s,
        };
        if nper == 0.0 {
            return CalcResult::new_error(Error::DIV, cell, "Division by 0".to_string());
        }
        CalcResult::Number(pv * rate * (per / nper - 1.0))
    }

    // RRI(nper, pv, fv)
    // Formula is
    // $$ \left(\frac{fv}{pv}\right)^{\frac{1}{nper}}-1  $$
    pub(crate) fn fn_rri(
        workbook: &Workbook,
        cells: &mut HashMap<(u32, i32, i32), CellState>,
        parsed_formulas: &Vec<Vec<Node>>,
        language: &Language,
        args: &[Node], cell: CellReferenceIndex) -> CalcResult {
        if args.len() != 3 {
            return CalcResult::new_args_number_error(cell);
        }
        let nper = match Model::get_number(workbook, cells, parsed_formulas, language, &args[0], cell) {
            Ok(f) => f,
            Err(s) => return s,
        };
        let pv = match Model::get_number(workbook, cells, parsed_formulas, language, &args[1], cell) {
            Ok(f) => f,
            Err(s) => return s,
        };
        let fv = match Model::get_number(workbook, cells, parsed_formulas, language, &args[2], cell) {
            Ok(f) => f,
            Err(s) => return s,
        };
        if nper <= 0.0 {
            return CalcResult::new_error(Error::NUM, cell, "nper should be >0".to_string());
        }
        if pv == 0.0 {
            // Note error is NUM not DIV/0 also bellow
            return CalcResult::new_error(Error::NUM, cell, "Division by 0".to_string());
        }
        let result = (fv / pv).powf(1.0 / nper) - 1.0;
        if result.is_infinite() {
            return CalcResult::new_error(Error::NUM, cell, "Division by 0".to_string());
        }
        if result.is_nan() {
            return CalcResult::new_error(Error::NUM, cell, "Invalid data for RRI".to_string());
        }

        CalcResult::Number(result)
    }

    // SLN(cost, salvage, life)
    // Formula is:
    // $$ \frac{cost-salvage}{life} $$
    pub(crate) fn fn_sln(
        workbook: &Workbook,
        cells: &mut HashMap<(u32, i32, i32), CellState>,
        parsed_formulas: &Vec<Vec<Node>>,
        language: &Language,
        args: &[Node], cell: CellReferenceIndex) -> CalcResult {
        if args.len() != 3 {
            return CalcResult::new_args_number_error(cell);
        }
        let cost = match Model::get_number(workbook, cells, parsed_formulas, language, &args[0], cell) {
            Ok(f) => f,
            Err(s) => return s,
        };
        let salvage = match Model::get_number(workbook, cells, parsed_formulas, language, &args[1], cell) {
            Ok(f) => f,
            Err(s) => return s,
        };
        let life = match Model::get_number(workbook, cells, parsed_formulas, language, &args[2], cell) {
            Ok(f) => f,
            Err(s) => return s,
        };
        if life == 0.0 {
            return CalcResult::new_error(Error::DIV, cell, "Division by 0".to_string());
        }
        let result = (cost - salvage) / life;

        CalcResult::Number(result)
    }

    // SYD(cost, salvage, life, per)
    // Formula is:
    // $$ \frac{(cost-salvage)*(life-per+1)*2}{life*(life+1)} $$
    pub(crate) fn fn_syd(
        workbook: &Workbook,
        cells: &mut HashMap<(u32, i32, i32), CellState>,
        parsed_formulas: &Vec<Vec<Node>>,
        language: &Language,
        args: &[Node], cell: CellReferenceIndex) -> CalcResult {
        if args.len() != 4 {
            return CalcResult::new_args_number_error(cell);
        }
        let cost = match Model::get_number(workbook, cells, parsed_formulas, language, &args[0], cell) {
            Ok(f) => f,
            Err(s) => return s,
        };
        let salvage = match Model::get_number(workbook, cells, parsed_formulas, language, &args[1], cell) {
            Ok(f) => f,
            Err(s) => return s,
        };
        let life = match Model::get_number(workbook, cells, parsed_formulas, language, &args[2], cell) {
            Ok(f) => f,
            Err(s) => return s,
        };
        let per = match Model::get_number(workbook, cells, parsed_formulas, language, &args[3], cell) {
            Ok(f) => f,
            Err(s) => return s,
        };
        if life == 0.0 {
            return CalcResult::new_error(Error::NUM, cell, "Division by 0".to_string());
        }
        if per > life || per <= 0.0 {
            return CalcResult::new_error(Error::NUM, cell, "per should be <= life".to_string());
        }
        let result = ((cost - salvage) * (life - per + 1.0) * 2.0) / (life * (life + 1.0));

        CalcResult::Number(result)
    }

    // NOMINAL(effective_rate, npery)
    // Formula is:
    // $$ n\times\left(\left(1+r\right)^{\frac{1}{n}}-1\right) $$
    // where:
    //   $r$ is the effective interest rate
    //   $n$ is the number of periods per year
    pub(crate) fn fn_nominal(
        workbook: &Workbook,
        cells: &mut HashMap<(u32, i32, i32), CellState>,
        parsed_formulas: &Vec<Vec<Node>>,
        language: &Language,
        args: &[Node], cell: CellReferenceIndex) -> CalcResult {
        if args.len() != 2 {
            return CalcResult::new_args_number_error(cell);
        }
        let effect_rate = match Model::get_number_no_bools(workbook, cells, parsed_formulas, language, &args[0], cell) {
            Ok(f) => f,
            Err(s) => return s,
        };
        let npery = match Model::get_number_no_bools(workbook, cells, parsed_formulas, language, &args[1], cell) {
            Ok(f) => f.floor(),
            Err(s) => return s,
        };
        if effect_rate <= 0.0 || npery < 1.0 {
            return CalcResult::new_error(Error::NUM, cell, "Invalid arguments".to_string());
        }
        let result = ((1.0 + effect_rate).powf(1.0 / npery) - 1.0) * npery;
        if result.is_infinite() {
            return CalcResult::new_error(Error::DIV, cell, "Division by 0".to_string());
        }
        if result.is_nan() {
            return CalcResult::new_error(Error::NUM, cell, "Invalid data for RRI".to_string());
        }

        CalcResult::Number(result)
    }

    // EFFECT(nominal_rate, npery)
    // Formula is:
    // $$ \left(1+\frac{r}{n}\right)^n-1 $$
    // where:
    //   $r$ is the nominal interest rate
    //   $n$ is the number of periods per year
    pub(crate) fn fn_effect(
        workbook: &Workbook,
        cells: &mut HashMap<(u32, i32, i32), CellState>,
        parsed_formulas: &Vec<Vec<Node>>,
        language: &Language,
        args: &[Node], cell: CellReferenceIndex) -> CalcResult {
        if args.len() != 2 {
            return CalcResult::new_args_number_error(cell);
        }
        let nominal_rate = match Model::get_number_no_bools(workbook, cells, parsed_formulas, language, &args[0], cell) {
            Ok(f) => f,
            Err(s) => return s,
        };
        let npery = match Model::get_number_no_bools(workbook, cells, parsed_formulas, language, &args[1], cell) {
            Ok(f) => f.floor(),
            Err(s) => return s,
        };
        if nominal_rate <= 0.0 || npery < 1.0 {
            return CalcResult::new_error(Error::NUM, cell, "Invalid arguments".to_string());
        }
        let result = (1.0 + nominal_rate / npery).powf(npery) - 1.0;
        if result.is_infinite() {
            return CalcResult::new_error(Error::DIV, cell, "Division by 0".to_string());
        }
        if result.is_nan() {
            return CalcResult::new_error(Error::NUM, cell, "Invalid data for RRI".to_string());
        }

        CalcResult::Number(result)
    }

    // PDURATION(rate, pv, fv)
    // Formula is:
    // $$ \frac{log(fv) - log(pv)}{log(1+r)} $$
    // where:
    //   * $r$ is the interest rate per period
    //   * $pv$ is the present value of the investment
    //   * $fv$ is the desired future value of the investment
    pub(crate) fn fn_pduration(
        workbook: &Workbook,
        cells: &mut HashMap<(u32, i32, i32), CellState>,
        parsed_formulas: &Vec<Vec<Node>>,
        language: &Language,
        args: &[Node], cell: CellReferenceIndex) -> CalcResult {
        if args.len() != 3 {
            return CalcResult::new_args_number_error(cell);
        }
        let rate = match Model::get_number(workbook, cells, parsed_formulas, language, &args[0], cell) {
            Ok(f) => f,
            Err(s) => return s,
        };
        let pv = match Model::get_number(workbook, cells, parsed_formulas, language, &args[1], cell) {
            Ok(f) => f,
            Err(s) => return s,
        };
        let fv = match Model::get_number(workbook, cells, parsed_formulas, language, &args[2], cell) {
            Ok(f) => f,
            Err(s) => return s,
        };
        if fv <= 0.0 || pv <= 0.0 || rate <= 0.0 {
            return CalcResult::new_error(Error::NUM, cell, "Invalid arguments".to_string());
        }
        let result = (fv.ln() - pv.ln()) / ((1.0 + rate).ln());
        if result.is_infinite() {
            return CalcResult::new_error(Error::DIV, cell, "Division by 0".to_string());
        }
        if result.is_nan() {
            return CalcResult::new_error(Error::NUM, cell, "Invalid data for RRI".to_string());
        }

        CalcResult::Number(result)
    }

    /// This next three functions deal with Treasure Bills or T-Bills for short
    /// They are zero-coupon that mature in one year or less.
    ///  Definitions:
    ///    $r$ be the discount rate
    ///    $v$ the face value of the Bill
    ///    $p$ the price of the Bill
    ///    $d_m$ is the number of days from the settlement to maturity
    /// Then:
    ///   $$ p = v \times\left(1-\frac{d_m}{r}\right) $$
    /// If d_m is less than 183 days the he Bond Equivalent Yield (BEY, here $y$) is given by:
    /// $$ y = \frac{F - B}{M}\times \frac{365}{d_m} = \frac{365\times r}{360-r\times d_m}
    /// If d_m>= 183 days things are a bit more complicated.
    /// Let $d_e = d_m - 365/2$ if $d_m <= 365$ or $d_e = 183$ if $d_m = 366$.
    /// $$ v = p\times \left(1+\frac{y}{2}\right)\left(1+d_e\times\frac{y}{365}\right) $$
    /// Together with the previous relation of $p$ and $v$ gives us a quadratic equation for $y$.

    // TBILLEQ(settlement, maturity, discount)
    pub(crate) fn fn_tbilleq(
        workbook: &Workbook,
        cells: &mut HashMap<(u32, i32, i32), CellState>,
        parsed_formulas: &Vec<Vec<Node>>,
        language: &Language,
        args: &[Node], cell: CellReferenceIndex) -> CalcResult {
        if args.len() != 3 {
            return CalcResult::new_args_number_error(cell);
        }
        let settlement = match Model::get_number_no_bools(workbook, cells, parsed_formulas, language, &args[0], cell) {
            Ok(f) => f,
            Err(s) => return s,
        };
        let maturity = match Model::get_number_no_bools(workbook, cells, parsed_formulas, language, &args[1], cell) {
            Ok(f) => f,
            Err(s) => return s,
        };
        let discount = match Model::get_number_no_bools(workbook, cells, parsed_formulas, language, &args[2], cell) {
            Ok(f) => f,
            Err(s) => return s,
        };
        if !is_valid_date(settlement) || !is_valid_date(maturity) {
            return CalcResult::new_error(Error::NUM, cell, "Invalid date".to_string());
        }
        if settlement > maturity {
            return CalcResult::new_error(
                Error::NUM,
                cell,
                "settlement should be <= maturity".to_string(),
            );
        }
        if !is_less_than_one_year(settlement as i64, maturity as i64) {
            return CalcResult::new_error(
                Error::NUM,
                cell,
                "maturity <= settlement + year".to_string(),
            );
        }
        if discount <= 0.0 {
            return CalcResult::new_error(Error::NUM, cell, "discount should be >0".to_string());
        }
        // days to maturity
        let d_m = maturity - settlement;
        let result = if d_m < 183.0 {
            365.0 * discount / (360.0 - discount * d_m)
        } else {
            // Equation here is:
            // (1-days*rate/360)*(1+y/2)*(1+d_extra*y/year)=1
            let year = if d_m == 366.0 { 366.0 } else { 365.0 };
            let d_extra = d_m - year / 2.0;
            let alpha = 1.0 - d_m * discount / 360.0;
            let beta = 0.5 + d_extra / year;
            // ay^2+by+c=0
            let a = d_extra * alpha / (year * 2.0);
            let b = alpha * beta;
            let c = alpha - 1.0;
            (-b + (b * b - 4.0 * a * c).sqrt()) / (2.0 * a)
        };
        if result.is_infinite() {
            return CalcResult::new_error(Error::DIV, cell, "Division by 0".to_string());
        }
        if result.is_nan() {
            return CalcResult::new_error(Error::NUM, cell, "Invalid data for RRI".to_string());
        }

        CalcResult::Number(result)
    }

    // TBILLPRICE(settlement, maturity, discount)
    pub(crate) fn fn_tbillprice(
        workbook: &Workbook,
        cells: &mut HashMap<(u32, i32, i32), CellState>,
        parsed_formulas: &Vec<Vec<Node>>,
        language: &Language,
        args: &[Node], cell: CellReferenceIndex) -> CalcResult {
        if args.len() != 3 {
            return CalcResult::new_args_number_error(cell);
        }
        let settlement = match Model::get_number_no_bools(workbook, cells, parsed_formulas, language, &args[0], cell) {
            Ok(f) => f,
            Err(s) => return s,
        };
        let maturity = match Model::get_number_no_bools(workbook, cells, parsed_formulas, language, &args[1], cell) {
            Ok(f) => f,
            Err(s) => return s,
        };
        let discount = match Model::get_number_no_bools(workbook, cells, parsed_formulas, language, &args[2], cell) {
            Ok(f) => f,
            Err(s) => return s,
        };
        if !is_valid_date(settlement) || !is_valid_date(maturity) {
            return CalcResult::new_error(Error::NUM, cell, "Invalid date".to_string());
        }
        if settlement > maturity {
            return CalcResult::new_error(
                Error::NUM,
                cell,
                "settlement should be <= maturity".to_string(),
            );
        }
        if !is_less_than_one_year(settlement as i64, maturity as i64) {
            return CalcResult::new_error(
                Error::NUM,
                cell,
                "maturity <= settlement + year".to_string(),
            );
        }
        if discount <= 0.0 {
            return CalcResult::new_error(Error::NUM, cell, "discount should be >0".to_string());
        }
        // days to maturity
        let d_m = maturity - settlement;
        let result = 100.0 * (1.0 - discount * d_m / 360.0);
        if result.is_infinite() {
            return CalcResult::new_error(Error::DIV, cell, "Division by 0".to_string());
        }
        if result.is_nan() || result < 0.0 {
            return CalcResult::new_error(Error::NUM, cell, "Invalid data for RRI".to_string());
        }

        CalcResult::Number(result)
    }

    // TBILLYIELD(settlement, maturity, pr)
    pub(crate) fn fn_tbillyield(
        workbook: &Workbook,
        cells: &mut HashMap<(u32, i32, i32), CellState>,
        parsed_formulas: &Vec<Vec<Node>>,
        language: &Language,
        args: &[Node], cell: CellReferenceIndex) -> CalcResult {
        if args.len() != 3 {
            return CalcResult::new_args_number_error(cell);
        }
        let settlement = match Model::get_number_no_bools(workbook, cells, parsed_formulas, language, &args[0], cell) {
            Ok(f) => f,
            Err(s) => return s,
        };
        let maturity = match Model::get_number_no_bools(workbook, cells, parsed_formulas, language, &args[1], cell) {
            Ok(f) => f,
            Err(s) => return s,
        };
        let pr = match Model::get_number_no_bools(workbook, cells, parsed_formulas, language, &args[2], cell) {
            Ok(f) => f,
            Err(s) => return s,
        };
        if !is_valid_date(settlement) || !is_valid_date(maturity) {
            return CalcResult::new_error(Error::NUM, cell, "Invalid date".to_string());
        }
        if settlement > maturity {
            return CalcResult::new_error(
                Error::NUM,
                cell,
                "settlement should be <= maturity".to_string(),
            );
        }
        if !is_less_than_one_year(settlement as i64, maturity as i64) {
            return CalcResult::new_error(
                Error::NUM,
                cell,
                "maturity <= settlement + year".to_string(),
            );
        }
        if pr <= 0.0 {
            return CalcResult::new_error(Error::NUM, cell, "discount should be >0".to_string());
        }
        let days = maturity - settlement;
        let result = (100.0 - pr) * 360.0 / (pr * days);

        CalcResult::Number(result)
    }

    // DOLLARDE(fractional_dollar, fraction)
    pub(crate) fn fn_dollarde(
        workbook: &Workbook,
        cells: &mut HashMap<(u32, i32, i32), CellState>,
        parsed_formulas: &Vec<Vec<Node>>,
        language: &Language,
        args: &[Node], cell: CellReferenceIndex) -> CalcResult {
        if args.len() != 2 {
            return CalcResult::new_args_number_error(cell);
        }
        let fractional_dollar = match Model::get_number_no_bools(workbook, cells, parsed_formulas, language, &args[0], cell) {
            Ok(f) => f,
            Err(s) => return s,
        };
        let mut fraction = match Model::get_number_no_bools(workbook, cells, parsed_formulas, language, &args[1], cell) {
            Ok(f) => f,
            Err(s) => return s,
        };
        if fraction < 0.0 {
            return CalcResult::new_error(Error::NUM, cell, "fraction should be >= 1".to_string());
        }
        if fraction < 1.0 {
            // this is not necessarily DIV/0
            return CalcResult::new_error(Error::DIV, cell, "fraction should be >= 1".to_string());
        }
        fraction = fraction.trunc();
        while fraction > 10.0 {
            fraction /= 10.0;
        }
        let t = fractional_dollar.trunc();
        let result = t + (fractional_dollar - t) * 10.0 / fraction;
        CalcResult::Number(result)
    }

    // DOLLARFR(decimal_dollar, fraction)
    pub(crate) fn fn_dollarfr(
        workbook: &Workbook,
        cells: &mut HashMap<(u32, i32, i32), CellState>,
        parsed_formulas: &Vec<Vec<Node>>,
        language: &Language,
        args: &[Node], cell: CellReferenceIndex) -> CalcResult {
        if args.len() != 2 {
            return CalcResult::new_args_number_error(cell);
        }
        let decimal_dollar = match Model::get_number_no_bools(workbook, cells, parsed_formulas, language, &args[0], cell) {
            Ok(f) => f,
            Err(s) => return s,
        };
        let mut fraction = match Model::get_number_no_bools(workbook, cells, parsed_formulas, language, &args[1], cell) {
            Ok(f) => f,
            Err(s) => return s,
        };
        if fraction < 0.0 {
            return CalcResult::new_error(Error::NUM, cell, "fraction should be >= 1".to_string());
        }
        if fraction < 1.0 {
            // this is not necessarily DIV/0
            return CalcResult::new_error(Error::DIV, cell, "fraction should be >= 1".to_string());
        }
        fraction = fraction.trunc();
        while fraction > 10.0 {
            fraction /= 10.0;
        }
        let t = decimal_dollar.trunc();
        let result = t + (decimal_dollar - t) * fraction / 10.0;
        CalcResult::Number(result)
    }

    // CUMIPMT(rate, nper, pv, start_period, end_period, type)
    pub(crate) fn fn_cumipmt(
        workbook: &Workbook,
        cells: &mut HashMap<(u32, i32, i32), CellState>,
        parsed_formulas: &Vec<Vec<Node>>,
        language: &Language,
        args: &[Node], cell: CellReferenceIndex) -> CalcResult {
        if args.len() != 6 {
            return CalcResult::new_args_number_error(cell);
        }
        let rate = match Model::get_number_no_bools(workbook, cells, parsed_formulas, language, &args[0], cell) {
            Ok(f) => f,
            Err(s) => return s,
        };
        let nper = match Model::get_number_no_bools(workbook, cells, parsed_formulas, language, &args[1], cell) {
            Ok(f) => f,
            Err(s) => return s,
        };
        let pv = match Model::get_number_no_bools(workbook, cells, parsed_formulas, language, &args[2], cell) {
            Ok(f) => f,
            Err(s) => return s,
        };
        let start_period = match Model::get_number_no_bools(workbook, cells, parsed_formulas, language, &args[3], cell) {
            Ok(f) => f.ceil() as i32,
            Err(s) => return s,
        };
        let end_period = match Model::get_number_no_bools(workbook, cells, parsed_formulas, language, &args[4], cell) {
            Ok(f) => f.trunc() as i32,
            Err(s) => return s,
        };
        // 0 at the end of the period, 1 at the beginning of the period
        let period_type = match Model::get_number_no_bools(workbook, cells, parsed_formulas, language, &args[5], cell) {
            Ok(f) => {
                if f == 0.0 {
                    false
                } else if f == 1.0 {
                    true
                } else {
                    return CalcResult::new_error(
                        Error::NUM,
                        cell,
                        "invalid period type".to_string(),
                    );
                }
            }
            Err(s) => return s,
        };
        if start_period > end_period {
            return CalcResult::new_error(
                Error::NUM,
                cell,
                "start period should come before end period".to_string(),
            );
        }
        if rate <= 0.0 || nper <= 0.0 || pv <= 0.0 || start_period < 1 {
            return CalcResult::new_error(Error::NUM, cell, "invalid parameters".to_string());
        }
        let mut result = 0.0;
        for period in start_period..=end_period {
            result += match compute_ipmt(rate, period as f64, nper, pv, 0.0, period_type) {
                Ok(f) => f,
                Err(error) => {
                    return CalcResult::Error {
                        error: error.0,
                        origin: cell,
                        message: error.1,
                    }
                }
            }
        }
        CalcResult::Number(result)
    }

    // CUMPRINC(rate, nper, pv, start_period, end_period, type)
    pub(crate) fn fn_cumprinc(
        workbook: &Workbook,
        cells: &mut HashMap<(u32, i32, i32), CellState>,
        parsed_formulas: &Vec<Vec<Node>>,
        language: &Language,
        args: &[Node], cell: CellReferenceIndex) -> CalcResult {
        if args.len() != 6 {
            return CalcResult::new_args_number_error(cell);
        }
        let rate = match Model::get_number_no_bools(workbook, cells, parsed_formulas, language, &args[0], cell) {
            Ok(f) => f,
            Err(s) => return s,
        };
        let nper = match Model::get_number_no_bools(workbook, cells, parsed_formulas, language, &args[1], cell) {
            Ok(f) => f,
            Err(s) => return s,
        };
        let pv = match Model::get_number_no_bools(workbook, cells, parsed_formulas, language, &args[2], cell) {
            Ok(f) => f,
            Err(s) => return s,
        };
        let start_period = match Model::get_number_no_bools(workbook, cells, parsed_formulas, language, &args[3], cell) {
            Ok(f) => f.ceil() as i32,
            Err(s) => return s,
        };
        let end_period = match Model::get_number_no_bools(workbook, cells, parsed_formulas, language, &args[4], cell) {
            Ok(f) => f.trunc() as i32,
            Err(s) => return s,
        };
        // 0 at the end of the period, 1 at the beginning of the period
        let period_type = match Model::get_number_no_bools(workbook, cells, parsed_formulas, language, &args[5], cell) {
            Ok(f) => {
                if f == 0.0 {
                    false
                } else if f == 1.0 {
                    true
                } else {
                    return CalcResult::new_error(
                        Error::NUM,
                        cell,
                        "invalid period type".to_string(),
                    );
                }
            }
            Err(s) => return s,
        };
        if start_period > end_period {
            return CalcResult::new_error(
                Error::NUM,
                cell,
                "start period should come before end period".to_string(),
            );
        }
        if rate <= 0.0 || nper <= 0.0 || pv <= 0.0 || start_period < 1 {
            return CalcResult::new_error(Error::NUM, cell, "invalid parameters".to_string());
        }
        let mut result = 0.0;
        for period in start_period..=end_period {
            result += match compute_ppmt(rate, period as f64, nper, pv, 0.0, period_type) {
                Ok(f) => f,
                Err(error) => {
                    return CalcResult::Error {
                        error: error.0,
                        origin: cell,
                        message: error.1,
                    }
                }
            }
        }
        CalcResult::Number(result)
    }

    // DDB(cost, salvage, life, period, [factor])
    pub(crate) fn fn_ddb(
        workbook: &Workbook,
        cells: &mut HashMap<(u32, i32, i32), CellState>,
        parsed_formulas: &Vec<Vec<Node>>,
        language: &Language,
        args: &[Node], cell: CellReferenceIndex) -> CalcResult {
        let arg_count = args.len();
        if !(4..=5).contains(&arg_count) {
            return CalcResult::new_args_number_error(cell);
        }
        let cost = match Model::get_number(workbook, cells, parsed_formulas, language, &args[0], cell) {
            Ok(f) => f,
            Err(s) => return s,
        };
        let salvage = match Model::get_number(workbook, cells, parsed_formulas, language, &args[1], cell) {
            Ok(f) => f,
            Err(s) => return s,
        };
        let life = match Model::get_number(workbook, cells, parsed_formulas, language, &args[2], cell) {
            Ok(f) => f,
            Err(s) => return s,
        };
        let period = match Model::get_number(workbook, cells, parsed_formulas, language, &args[3], cell) {
            Ok(f) => f,
            Err(s) => return s,
        };
        // The rate at which the balance declines.
        let factor = if arg_count > 4 {
            match Model::get_number_no_bools(workbook, cells, parsed_formulas, language, &args[4], cell) {
                Ok(f) => f,
                Err(s) => return s,
            }
        } else {
            // If factor is omitted, it is assumed to be 2 (the double-declining balance method).
            2.0
        };
        if period > life || cost < 0.0 || salvage < 0.0 || period <= 0.0 || factor <= 0.0 {
            return CalcResult::new_error(Error::NUM, cell, "invalid parameters".to_string());
        };
        // let period_trunc = period.floor() as i32;
        let mut rate = factor / life;
        if rate > 1.0 {
            rate = 1.0
        };
        let value = if rate == 1.0 {
            if period == 1.0 {
                cost
            } else {
                0.0
            }
        } else {
            cost * (1.0 - rate).powf(period - 1.0)
        };
        let new_value = cost * (1.0 - rate).powf(period);
        let result = f64::max(value - f64::max(salvage, new_value), 0.0);
        CalcResult::Number(result)
    }

    // DB(cost, salvage, life, period, [month])
    pub(crate) fn fn_db(
        workbook: &Workbook,
        cells: &mut HashMap<(u32, i32, i32), CellState>,
        parsed_formulas: &Vec<Vec<Node>>,
        language: &Language,
        args: &[Node], cell: CellReferenceIndex) -> CalcResult {
        let arg_count = args.len();
        if !(4..=5).contains(&arg_count) {
            return CalcResult::new_args_number_error(cell);
        }
        let cost = match Model::get_number(workbook, cells, parsed_formulas, language, &args[0], cell) {
            Ok(f) => f,
            Err(s) => return s,
        };
        let salvage = match Model::get_number(workbook, cells, parsed_formulas, language, &args[1], cell) {
            Ok(f) => f,
            Err(s) => return s,
        };
        let life = match Model::get_number(workbook, cells, parsed_formulas, language, &args[2], cell) {
            Ok(f) => f,
            Err(s) => return s,
        };
        let period = match Model::get_number(workbook, cells, parsed_formulas, language, &args[3], cell) {
            Ok(f) => f,
            Err(s) => return s,
        };
        let month = if arg_count > 4 {
            match Model::get_number_no_bools(workbook, cells, parsed_formulas, language, &args[4], cell) {
                Ok(f) => f.trunc(),
                Err(s) => return s,
            }
        } else {
            12.0
        };
        if month == 12.0 && period > life
            || (period > life + 1.0)
            || month <= 0.0
            || month > 12.0
            || period <= 0.0
            || cost < 0.0
        {
            return CalcResult::new_error(Error::NUM, cell, "invalid parameters".to_string());
        };
        if cost == 0.0 {
            return CalcResult::Number(0.0);
        }
        // rounded to three decimal places
        // FIXME: We should have utilities for this (see to_precision)
        let rate = f64::round((1.0 - f64::powf(salvage / cost, 1.0 / life)) * 1000.0) / 1000.0;

        let mut result = cost * rate * month / 12.0;

        let period = period.floor() as i32;
        let life = life.floor() as i32;

        // Depreciation for the first and last periods is a special case.
        if period == 1 {
            return CalcResult::Number(result);
        };

        for _ in 0..period - 2 {
            result += (cost - result) * rate;
        }

        if period == life + 1 {
            // last period
            return CalcResult::Number((cost - result) * rate * (12.0 - month) / 12.0);
        }

        CalcResult::Number(rate * (cost - result))
    }
}
