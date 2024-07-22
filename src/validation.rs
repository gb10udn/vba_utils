use calamine::{Reader, open_workbook, Xlsx};
use std::collections::HashSet;
use regex::Regex;


pub fn validate_no_duplicated_procedure_names(path: &str) -> Vec<String> {
    let mut workbook: Xlsx<_> = open_workbook(path).expect("Cannot open file");
    let mut procedure_names: Vec<String> = vec![];
    if let Some(Ok(mut vba)) = workbook.vba_project() {
        let vba = vba.to_mut();
        let module_names = vba.get_module_names();

        for module_name in module_names {
            let vba_code = vba.get_module(module_name).unwrap();
            for proc in extract_procedure_names(&vba_code) {
                procedure_names.push(proc);
            }
        }
    }
    extract_duplicate_procedure_names(&procedure_names)
}

/// プロシージャ名を抜き取る関数。
fn extract_procedure_names(code: &String) -> Vec<String> {
    let mut result = vec![];
    let re = Regex::new(r"(?i)^\s*(Public|Private)?\s*(Sub|Function)\s+([A-Za-z0-9_]+)").unwrap();
    for one_line in code.split("\n") {
        if let Some(caps) = re.captures(one_line) {
            if let Some(proc_name) = caps.get(3) {
                result.push(proc_name.as_str().to_string());
            }
        }
    }
    result
}

fn extract_duplicate_procedure_names(vec: &Vec<String>) -> Vec<String> {
    let mut result: Vec<String> = vec![];
    let mut set = HashSet::new();
    for item in vec {
        if !set.insert(item) {
            result.push(item.clone());
        }
    }
    result
}