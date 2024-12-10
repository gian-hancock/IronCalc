#![allow(clippy::unwrap_used)]
#![allow(clippy::panic)]

//! Tests an Excel xlsx file.
//! Returns a list of differences in json format.
//! Saves an IronCalc version
//! This is primary for QA internal testing and will be superseded by an official
//! IronCalc CLI.
//!
//! Usage: test file.xlsx

use std::path;

use ironcalc::{export::{save_to_icalc, save_to_xlsx}, import::load_from_xlsx};
use ironcalc_base::{expressions::parser::Node, Model};

fn main_old() {
    let args: Vec<_> = std::env::args().collect();
    if args.len() != 2 {
        panic!("Usage: {} <file.xlsx>", args[0]);
    }
    // first test the file
    let file_name = &args[1];

    let file_path = path::Path::new(file_name);
    let base_name = file_path.file_stem().unwrap().to_str().unwrap();
    let output_file_name = &format!("{base_name}.ic");
    let model = load_from_xlsx(file_name, "en", "UTC").unwrap();
    save_to_icalc(&model, output_file_name).unwrap();
}

fn main() {
    let args: Vec<_> = std::env::args().collect();
    if args.len() != 2 {
        println!("Usage: {} <directory_path>", args[0]);
        panic!();
    }

    let mut count_by_fn = std::collections::HashMap::new();
    let mut output_model = Model::new_empty("formulas-and-errors", "en", "UTC").unwrap();
    let mut row = 1;

    // first test the file
    let dir_path = &args[1];

    let xlsx_files = get_xlsx_files_in_dir(path::Path::new(dir_path)).unwrap();
    for file_path in xlsx_files {
        println!("===== Forumlas in: {} =====", file_path.display());
        let model = if let Ok(model) = load_from_xlsx(file_path.to_str().unwrap(), "en", "UTC") {
            model
        } else {
            println!("Failed to load model from: {}", file_path.display());
            continue;
        };
        // ===== Find all functions in the model ===== //
        let functions = model.parsed_formulas.iter().flat_map(|f| f.iter())
            .filter_map(|f| match f {
            Node::FunctionKind{ kind, .. } => Some(kind.clone()),
            _ => None,
        });

        for function in functions {
            output_model.update_cell_with_text(0, row, 1, &format!("{}", &function)).unwrap();
            output_model.update_cell_with_text(0, row, 2, &format!("{}", &file_path.display())).unwrap();
            row += 1;
            let count = count_by_fn.get(&function).unwrap_or(&0) + 1;
            count_by_fn.insert(function, count);
        }
        dbg!(&count_by_fn);
        count_by_fn.clear();
        // dbg!(&functions.collect::<Vec<_>>());
    }
    save_to_xlsx(&output_model, "function counts.xlsx").unwrap();
}

fn get_xlsx_files_in_dir(dir: &path::Path) -> std::io::Result<impl Iterator<Item = std::path::PathBuf>> {
    let mut paths = Vec::new();
    if dir.is_dir() {
        for entry in std::fs::read_dir(dir)? {
            let entry = entry?;
            let path = entry.path();
            if path.is_dir() {
                paths.extend(get_xlsx_files_in_dir(&path)?);
            } else {
                let file_name = path.file_name().unwrap().to_str().unwrap();
                if file_name.ends_with(".xlsx") {
                    paths.push(path);
                }
            }
        }
    }
    Ok(paths.into_iter())
}