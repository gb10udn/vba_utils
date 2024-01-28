use calamine::{Reader, open_workbook, Xlsx};
use std::fs::File;
use std::io::Write;
use std::io;

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

        for module_name in module_names {
            let vba_code = vba.get_module(module_name).unwrap();
            for one_line in vba_code.split("\n") {
                // TODO: 240128 テストコードの関数を書き出さないようにする。
                // TODO: 240128 モジュール名を関数名の先頭につける
                // TODO: 240128 docstring 以外のコメントをすべて削除する？
                if one_line.starts_with("Option Explicit") == false && one_line.starts_with("Attribute ") == false {
                    summary_vba_code.push_str(one_line);
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