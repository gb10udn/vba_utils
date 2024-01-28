use calamine::{Reader, open_workbook, Xlsx};
use std::{fs::File, io::Write, io};
use regex::Regex;

fn main() {
    let path = "./vba_utils.xlsm";
    write_each_code(path);
    write_summary_code(path, "Utils", true, true);
    stop();
}

/// 各モジュールを .bas ファイルとして保存する関数。
fn write_each_code(path: &str) {  // TODO: 240128 エクセルファイル名で、別のフォルダ名をつけるとかのするかどうかを選択できるといいかも？
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
fn write_summary_code(path: &str, module_name: &str, remove_test_code: bool, add_module_name: bool) {  // FIXME: 240128 Utils.bas の名称が衝突する場合のエラー処理を書くこと。 (panic! で強制的にストップさせてよいかも？)
    let mut workbook: Xlsx<_> = open_workbook(path).expect("Cannot open file");
    
    if let Some(Ok(mut vba)) = workbook.vba_project() {
        let vba = vba.to_mut();
        let module_names = vba.get_module_names();
        
        let header = format!("Attribute VB_Name = \"{}\"\nOption Explicit\n", module_name);
        let mut summary_vba_code = String::from(&header);
        
        let re_test_block_start = Regex::new(r"Function TEST_|Sub TEST_").unwrap();
        let re_test_block_end = Regex::new(r"End Function|End Sub").unwrap();
        
        for module_name in module_names {
            let vba_code = vba.get_module(module_name).unwrap();
            let mut is_test_block = false;
            for one_line in vba_code.split("\n") {
                // TODO: 240128 docstring 以外のコメントをすべて削除する？
                if (remove_test_code == true) && (is_test_block == false) && (re_test_block_start.is_match(one_line) == true) {
                    is_test_block = true;
                }
                
                if (is_test_block == false) && (one_line.starts_with("Option Explicit") == false) && (one_line.starts_with("Attribute ") == false) {
                    let editted_one_line;
                    
                    if add_module_name == true && one_line.contains("Function") == true && one_line.contains("End") == false {  // FIXME: 240128 not DRY
                        let word_vec: Vec<&str> = one_line.split("Function").collect();
                        editted_one_line = format!("{}Function {}_{}", word_vec[0], module_name, word_vec[1].trim());
                    } else if add_module_name == true && one_line.contains("Sub") == true && one_line.contains("End") == false {
                        let word_vec: Vec<&str> = one_line.split("Sub").collect();
                        editted_one_line = format!("{}Sub {}_{}", word_vec[0], module_name, word_vec[1].trim());
                    } else {
                        editted_one_line = String::from(one_line);  // FIXME: 240128 editted と言いながら、そのまま使っているのが微妙かも。    
                    }

                    summary_vba_code.push_str(&format!("{}\n", editted_one_line));
                }

                if (remove_test_code == true) && (is_test_block == true) && (re_test_block_end.is_match(one_line) == true) {
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