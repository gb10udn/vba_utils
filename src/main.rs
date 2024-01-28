use calamine::{Reader, open_workbook, Xlsx};
use std::{fs::File, io::Write, io};
use regex::Regex;

fn main() {
    let path = "./vba_utils.xlsm";
    write_each_code(path);
    write_summary_code(path, "utils");
    stop();
}

/// 各モジュールを .bas ファイルとして保存する関数。
fn write_each_code(path: &str) {
    let mut workbook: Xlsx<_> = open_workbook(path).expect("Cannot open file");
    
    if let Some(Ok(mut vba)) = workbook.vba_project() {
        let vba = vba.to_mut();
        let module_names = vba.get_module_names();

        for module_name in module_names {
            let vba_code = vba.get_module(module_name).unwrap();
            write_text(&vba_code, &format!("{}.bas", module_name)).unwrap();
        }
    }
}

/// 複数のモジュールを結合して、１つの .bas として保存する。
fn write_summary_code(path: &str, module_name: &str) {
    let mut workbook: Xlsx<_> = open_workbook(path).expect("Cannot open file");
    
    if let Some(Ok(mut vba)) = workbook.vba_project() {
        let vba = vba.to_mut();
        let module_names = vba.get_module_names();
        
        let header = format!("Attribute VB_Name = \"{}\"\nOption Explicit\n", module_name);
        let mut summary_vba_code = String::from(&header);
        let re_test_start = Regex::new(r"Function TEST_|Sub TEST_").unwrap();
        let re_test_end = Regex::new(r"End Function|End Sub").unwrap();
        for module_name in module_names {
            let vba_code = vba.get_module(module_name).unwrap();
            let mut is_test_block = false;
            for one_line in vba_code.split("\n") {
                // TODO: 240128 モジュール名を関数名の先頭につける
                // TODO: 240128 docstring 以外のコメントをすべて削除する？
                if (is_test_block == false) && (re_test_start.is_match(one_line) == true) {
                    is_test_block = true;
                }

                if (is_test_block == false) && (one_line.starts_with("Option Explicit") == false) && (one_line.starts_with("Attribute ") == false) {
                    summary_vba_code.push_str(&format!("{}\n", one_line));
                }

                if (is_test_block == true) && (re_test_end.is_match(one_line) == true) {
                    is_test_block = false;
                }
            }
        }
        write_text(&summary_vba_code, "utils.bas").unwrap();
    }
}

fn write_text(text: &str, dst: &str) -> Result<(), Box<dyn std::error::Error>> {
    let mut file = File::create(dst)?;
    write!(file, "{}", text)?;
    Ok(())
}

fn stop() {
    println!("finished !!! Please input enter key");
    let mut a = String::new();
    let _  = io::stdin().read_line(&mut a).expect("");
}