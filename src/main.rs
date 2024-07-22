mod validation;
mod summary;

use calamine::{Reader, open_workbook, Xlsx};
use std::{fs::File, io::Write, io};

fn main() {
    let path = "./vba_utils.xlsm";
    let duplicated_procedure_names = validation::validate_no_duplicated_procedure_names(path);

    if duplicated_procedure_names.len() > 0 {
        panic!("DuplicateProcedureNameError ... -> {:?}", duplicated_procedure_names);
    }
    write_each_code(path);

    let dst_module_name = "Utils";
    if let Ok(code) = summary::summarize(path, dst_module_name, true) {
        write_text(&code, &format!("{}.bas", dst_module_name)).unwrap();
    } else {
        panic!("Fail to summarize ....");
    }
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