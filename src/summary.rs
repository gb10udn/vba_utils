use calamine::{Reader, open_workbook, Xlsx, Error};
use regex::Regex;

/// 複数のモジュールを結合して、１つの .bas として保存する関数。
pub fn summarize(path: &str, dst_module_name: &str, remove_test_code: bool) -> Result<String, Error> {  // FIXME: 240128 Utils.bas の名称が衝突する場合のエラー処理を書くこと。 (panic! で強制的にストップさせてよいかも？)
    let mut workbook: Xlsx<_> = open_workbook(path).expect("Cannot open file");
    
    if let Some(Ok(mut vba)) = workbook.vba_project() {
        let vba = vba.to_mut();
        let module_names = vba.get_module_names();
        
        let header = format!("Attribute VB_Name = \"{}\"\nOption Explicit\n", dst_module_name);
        let mut summary_vba_code = String::from(&header);
        
        let re_test_block_start = Regex::new(r"Function TEST_|Sub TEST_").unwrap();
        let re_test_block_end = Regex::new(r"End Function|End Sub").unwrap();
        
        for module_name in &module_names {
            let vba_code = vba.get_module(module_name).unwrap();
            let mut is_test_block = false;
            for one_line in vba_code.split("\n") {
                // TODO: 240128 docstring 以外のコメントをすべて削除する？
                if (remove_test_code == true) && (is_test_block == false) && (re_test_block_start.is_match(one_line) == true) {
                    is_test_block = true;
                }
                
                if (is_test_block == false) && (one_line.starts_with("Option Explicit") == false) && (one_line.starts_with("Attribute ") == false) {
                    summary_vba_code.push_str(&format!("{}\n", remove_qualifier(one_line, &module_names)));
                }

                if (remove_test_code == true) && (is_test_block == true) && (re_test_block_end.is_match(one_line) == true) {
                    is_test_block = false;
                }
            }
        }
        Ok(summary_vba_code)
    } else {
        Err(From::from("expected at least one record but got none"))
    }
}


/// 限定子 (qualifier) を削除する。{モジュール名}.{関数名} のように呼び出す場合の、{モジュール名}. の部分を削除する。
fn remove_qualifier(one_line: &str, module_names: &Vec<&str>) -> String {
    let mut result = one_line.to_string();
    for module_name in module_names {
        if module_name != &"ThisWorkbook" {
            let target = format!("{}.", module_name);
            result = result.replace(&target, "");
        }
    }
    result
}