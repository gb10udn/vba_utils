use calamine::{Reader, open_workbook, Xlsx};
use std::fs::File;
use std::io::{Write};
use std::io;

fn main() {
    let path = "./vba_utils.xlsm";
    let mut workbook: Xlsx<_> = open_workbook(path).expect("Cannot open file");
    
    if let Some(Ok(mut vba)) = workbook.vba_project() {
        let vba = vba.to_mut();
        let module_names = vba.get_module_names();

        for module_name in module_names {
            let vba_code = vba.get_module(module_name).unwrap();  // FIXME: 240127 エラーハンドリングを修正せよ。
            write_text(&vba_code, &format!("{}.bas", module_name)).unwrap();  // TODO: 240127 utils.bas にまとめる、minify する機能を追加する。
        }
    }
    stop();
}

fn write_text(text: &str, dst: &str) -> Result<(), Box<dyn std::error::Error>> {  // FIXME: 240127 エラーハンドリングを修正せよ。
    let mut file = File::create(dst)?;
    write!(file, "{}", text)?;
    Ok(())
}

fn stop() {
    println!("finished !!! Please input enter key");
    let mut a = String::new();
    let _  = io::stdin().read_line(&mut a).expect("");
}