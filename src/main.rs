use chrono::Local;
use rfd::FileDialog;
use rust_xlsxwriter::Workbook;
use serde_json::Value;
use std::io::{self, BufReader, Write};
use std::{env, fs::File};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // 提示用户选择 zh-cn.json 文件
    let file_path = FileDialog::new()
        .set_title("选择 zh-cn.json 文件")
        .add_filter("JSON 文件", &["json"])
        .pick_file();

    let json_file_path = match file_path {
        Some(path) => path,
        None => {
            println!("未选择文件，程序退出。");
            return Ok(());
        }
    };

    let mut input = String::new();
    println!("请输入一个字符串（默认值为 'i18n.'），然后按回车继续...");

    // 读取用户输入
    io::stdout().flush()?; // 确保输出被刷新
    io::stdin().read_line(&mut input)?;

    // 去除输入字符串末尾的换行符
    let prefix = input.trim();
    let prefix = if prefix.is_empty() { "i18n." } else { prefix };

    // 读取 JSON 文件
    let current_dir = env::current_dir()?;
    
    let file = File::open(json_file_path)?;
    let reader = BufReader::new(file);
    let json: Value = serde_json::from_reader(reader)?;

    // 获取 "i18n" 对象，并确保其为有序的 Map
    let json_object = json.get("i18n").ok_or("i18n not found")?;
    let json_map = json_object.as_object().ok_or("i18n is not an object")?;

    // 创建 Excel 文件
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    // 写入数据到 Excel 工作表
    let mut row_index = 0;
    for (key, value) in json_map.iter() {
        let key_str = format!("{}{}", prefix, key);
        let value_str = value.as_str().unwrap_or("");

        worksheet.write_string(row_index, 0, &key_str)?;
        worksheet.write_string(row_index, 1, value_str)?;
        row_index += 1;
    }

    let now = Local::now();
    let timestamp = now.format("%Y%m%d%H%M%S").to_string();
    let output_file_name = format!("zh-cn_{}.xlsx", timestamp);
    let output_file_path = current_dir.join(output_file_name);

    // 保存 Excel 文件
    workbook.save(output_file_path)?;

    Ok(())
}
