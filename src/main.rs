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

/// 複数のモジュールを結合して、１つの .bas として保存する関数。
fn write_summary_code(path: &str, dst_module_name: &str, remove_test_code: bool, with_module_name: bool) {  // FIXME: 240128 Utils.bas の名称が衝突する場合のエラー処理を書くこと。 (panic! で強制的にストップさせてよいかも？)
    let mut workbook: Xlsx<_> = open_workbook(path).expect("Cannot open file");
    
    if let Some(Ok(mut vba)) = workbook.vba_project() {
        let vba = vba.to_mut();
        let module_names = vba.get_module_names();
        
        let header = format!("Attribute VB_Name = \"{}\"\nOption Explicit\n", dst_module_name);
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
                    if with_module_name {
                        summary_vba_code.push_str(&format!("{}\n", add_module_name(one_line, module_name)));
                    } else {
                        summary_vba_code.push_str(&format!("{}\n", one_line));
                    }
                }

                if (remove_test_code == true) && (is_test_block == true) && (re_test_block_end.is_match(one_line) == true) {
                    is_test_block = false;
                }
            }
        }
        write_text(&summary_vba_code, &format!("{}.bas", dst_module_name)).unwrap();
    }
}

fn write_text(text: &str, dst: &str) -> Result<(), Box<dyn std::error::Error>> {
    let mut file = File::create(dst)?;
    write!(file, "{}", text)?;
    Ok(())
}

/// Sub or Function の名前の先頭にモジュール名を付与する関数
/// Ex. Sub Hoge() --> Sub Module_Hoge() (Hoge というサブプロシージャの先頭に、モジュール名の Module を追加して返した。)
fn add_module_name(arg: &str, module_name: &str) -> String {
    for target_reserved_word in vec!["Public Sub", "Public Function"] {
        if arg.contains(target_reserved_word) == true && arg.contains("End") == false {
            let word_vec: Vec<&str> = arg.split(target_reserved_word).collect();
            return format!("{}{} {}_{}", word_vec[0], target_reserved_word, module_name, word_vec[1].trim());
        }
    }
    return String::from(arg);
}

fn stop() {
    println!("finished !!! Please input enter key");
    let mut a = String::new();
    let _  = io::stdin().read_line(&mut a).expect("");
}