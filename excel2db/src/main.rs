use std::env;
use std::fs::File;
use std::io::{self, Write};
use std::path::{Path, PathBuf};

use clap::{Arg, Command};
use rusqlite::{Connection, Result as SqliteResult};
use calamine::{Reader, Xlsx, Xls, open_workbook};
use csv;

#[derive(Debug)]
enum FileType {
    Xlsx,
    Xls,
    Csv,
}

impl FileType {
    fn from_extension(path: &Path) -> Option<Self> {
        match path.extension()?.to_str()?.to_lowercase().as_str() {
            "xlsx" => Some(FileType::Xlsx),
            "xls" => Some(FileType::Xls),
            "csv" => Some(FileType::Csv),
            _ => None,
        }
    }
}

fn main() {
    let args: Vec<String> = env::args().collect();

    if args.len() > 1 && !args[1].starts_with('-') {
        // 拖拽多文件批量处理
        for file_path in &args[1..] {
            if let Err(e) = process_file_drag_drop(file_path) {
                eprintln!("处理文件 {} 时出错: {}", file_path, e);
            }
        }
        return;
    }

    let matches = Command::new("Excel/CSV to SQLite Converter")
        .version("1.1")
        .author("ZYG")
        .about("将Excel和CSV文件转换为SQLite数据库")
        .arg(
            Arg::new("path")
                .short('p')
                .long("path")
                .value_name("FILE")
                .help("指定要转换的文件路径")
                .required(false)
        )
        .get_matches();

    if let Some(file_path) = matches.get_one::<String>("path") {
        if let Err(e) = process_file_interactive(file_path) {
            eprintln!("处理文件时出错: {}", e);
        }
    } else {
        interactive_mode();
    }
}

fn process_file_drag_drop(file_path: &str) -> Result<(), Box<dyn std::error::Error>> {
    let path = Path::new(file_path);

    if !path.exists() {
        return Err(format!("文件不存在: {}", file_path).into());
    }

    let file_type = FileType::from_extension(path)
        .ok_or_else(|| format!("不支持的文件类型: {}", file_path))?;

    // Excel 读取工作表
    if let FileType::Xlsx | FileType::Xls = file_type {
        let sheet_names = get_excel_sheet_names(path, &file_type)?;
        let (output_path, sheet_name) = if sheet_names.len() > 1 {
            // 多表默认选第一个表（拖拽场景不交互）
            let sheet_name = &sheet_names[0];
            let output_path = make_output_path_with_sheet(path, sheet_name);
            (output_path, Some(sheet_name.clone()))
        } else {
            (path.with_extension("db"), None)
        };

        convert_file(path, &output_path, &file_type, sheet_name.as_deref(), None)?;
        println!("文件 {} 转换完成，输出: {}", file_path, output_path.display());
    } else {
        // CSV
        let output_path = path.with_extension("db");
        convert_file(path, &output_path, &file_type, None, None)?;
        println!("文件 {} 转换完成，输出: {}", file_path, output_path.display());
    }

    Ok(())
}

fn process_file_interactive(file_path: &str) -> Result<(), Box<dyn std::error::Error>> {
    let path = Path::new(file_path);

    if !path.exists() {
        return Err(format!("文件不存在: {}", file_path).into());
    }

    let file_type = FileType::from_extension(path)
        .ok_or_else(|| format!("不支持的文件类型: {}", file_path))?;

    if let FileType::Xlsx | FileType::Xls = file_type {
        let sheet_names = get_excel_sheet_names(path, &file_type)?;

        // 多表选择
        let (selected_sheet, output_path) = if sheet_names.len() > 1 {
            let idx = select_worksheet(&sheet_names)?;
            let sheet_name = &sheet_names[idx];
            (Some(sheet_name.clone()), make_output_path_with_sheet(path, sheet_name))
        } else {
            // 单表
            (None, path.with_extension("db"))
        };

        let headers = if let Some(sheet) = &selected_sheet {
            read_excel_headers_by_name(path, &file_type, sheet)?
        } else {
            read_headers(path, &file_type)?
        };

        println!("检测到的表头:");
        for (i, header) in headers.iter().enumerate() {
            println!("  {}: {}", i + 1, header);
        }

        print!("是否确认使用这些表头? (y/n): ");
        io::stdout().flush()?;
        
        let mut input = String::new();
        io::stdin().read_line(&mut input)?;
        if input.trim().to_lowercase() != "y" {
            println!("转换已取消");
            return Ok(());
        }
        println!("正在转换，请稍后...");

        convert_file(path, &output_path, &file_type, selected_sheet.as_deref(), Some(&headers))?;
        println!("转换完成，输出文件: {}", output_path.display());
    } else {
        // CSV处理
        let headers = read_headers(path, &file_type)?;

        println!("检测到的表头:");
        for (i, header) in headers.iter().enumerate() {
            println!("  {}: {}", i + 1, header);
        }

        print!("是否确认使用这些表头? (y/n): ");
        io::stdout().flush()?;
        
        
        let mut input = String::new();
        io::stdin().read_line(&mut input)?;
        if input.trim().to_lowercase() != "y" {
            println!("转换已取消");
            return Ok(());
        }
        
        let output_path = path.with_extension("db");
        println!("正在转换，请稍后...");
        convert_file(path, &output_path, &file_type, None, Some(&headers))?;
        println!("转换完成，输出文件: {}", output_path.display());
    }

    Ok(())
}

fn interactive_mode() {
    println!("Excel/CSV to SQLite 转换器");
    println!("请输入要转换的文件路径（右键文件复制地址需要去掉引号）:");

    let mut input = String::new();
    if io::stdin().read_line(&mut input).is_ok() {
        let file_path = input.trim();
        if let Err(e) = process_file_interactive(file_path) {
            eprintln!("处理文件时出错: {}", e);
        }
    }

    println!("按回车键退出...");
    let mut _input = String::new();
    let _ = io::stdin().read_line(&mut _input);
}

/// 获取Excel所有工作表名
fn get_excel_sheet_names(
    path: &Path,
    file_type: &FileType,
) -> Result<Vec<String>, Box<dyn std::error::Error>> {
    match file_type {
        FileType::Xlsx => {
            let workbook: Xlsx<_> = open_workbook(path)?;
            Ok(workbook.sheet_names().to_owned())
        }
        FileType::Xls => {
            let workbook: Xls<_> = open_workbook(path)?;
            Ok(workbook.sheet_names().to_owned())
        }
        FileType::Csv => Err("CSV没有工作表列表".into()),
    }
}

/// 选择工作表索引
fn select_worksheet(sheet_names: &[String]) -> Result<usize, Box<dyn std::error::Error>> {
    println!("找到多个工作表，请输入要导出的工作表索引（0 - {}）：", sheet_names.len() - 1);
    for (i, name) in sheet_names.iter().enumerate() {
        println!("  [{}] {}", i, name);
    }

    loop {
        print!("输入索引: ");
        io::stdout().flush()?;
        let mut input = String::new();
        io::stdin().read_line(&mut input)?;
        if let Ok(idx) = input.trim().parse::<usize>() {
            if idx < sheet_names.len() {
                return Ok(idx);
            }
        }
        println!("输入无效，请重新输入");
    }
}

/// 根据原文件路径和表名生成输出文件路径
fn make_output_path_with_sheet(path: &Path, sheet_name: &str) -> PathBuf {
    let stem = path.file_stem().and_then(|s| s.to_str()).unwrap_or("output");
    let safe_sheet_name = sheet_name.replace('/', "_").replace('\\', "_");
    let new_name = format!("{} - {}.db", stem, safe_sheet_name);
    path.with_file_name(new_name)
}

/// 读取Excel特定工作表表头
fn read_excel_headers_by_name(
    path: &Path,
    file_type: &FileType,
    sheet_name: &str,
) -> Result<Vec<String>, Box<dyn std::error::Error>> {
    match file_type {
        FileType::Xlsx => read_excel_headers::<Xlsx<_>>(path, Some(sheet_name)),
        FileType::Xls => read_excel_headers::<Xls<_>>(path, Some(sheet_name)),
        FileType::Csv => Err("CSV没有工作表".into()),
    }
}

/// 读取Excel表头，sheet_name为None时取第一个表
fn read_excel_headers<R>(
    path: &Path,
    sheet_name: Option<&str>,
) -> Result<Vec<String>, Box<dyn std::error::Error>>
where
    R: Reader<std::io::BufReader<std::fs::File>>,
    <R as Reader<std::io::BufReader<std::fs::File>>>::Error: std::error::Error + std::fmt::Display + 'static,
{
    let mut workbook: R = open_workbook(path)?;

    let sheet_names = workbook.sheet_names();
    let sheet_name = if let Some(name) = sheet_name {
        name.to_string()
    } else {
        sheet_names.get(0)
            .ok_or_else(|| "Excel没有工作表".to_string())?
            .to_string()
    };

    let sheet = workbook
        .worksheet_range(&sheet_name)
        .ok_or_else(|| format!("工作表 '{}' 不存在", sheet_name))?;

    let range = sheet?;

    if let Some(first_row) = range.rows().next() {
        let headers = first_row.iter().map(|cell| cell.to_string()).collect();
        Ok(headers)
    } else {
        Err("工作表为空，找不到表头".into())
    }
}

/// 读取表头，支持Excel和CSV
fn read_headers(path: &Path, file_type: &FileType) -> Result<Vec<String>, Box<dyn std::error::Error>> {
    match file_type {
        FileType::Xlsx => read_excel_headers::<Xlsx<_>>(path, None),
        FileType::Xls => read_excel_headers::<Xls<_>>(path, None),
        FileType::Csv => read_csv_headers(path),
    }
}

/// 读取CSV表头
fn read_csv_headers(path: &Path) -> Result<Vec<String>, Box<dyn std::error::Error>> {
    let file = File::open(path)?;
    let mut reader = csv::Reader::from_reader(file);
    let headers: Vec<String> = reader.headers()?.iter().map(|h| h.to_string()).collect();
    Ok(headers)
}

/// 通用转换入口
fn convert_file(
    input_path: &Path,
    output_path: &Path,
    file_type: &FileType,
    sheet_name: Option<&str>,
    headers: Option<&Vec<String>>,
) -> Result<(), Box<dyn std::error::Error>> {
    let mut conn = Connection::open(output_path)?;

    match file_type {
        FileType::Xlsx => convert_excel::<Xlsx<_>>(&mut conn, input_path, sheet_name, headers)?,
        FileType::Xls => convert_excel::<Xls<_>>(&mut conn, input_path, sheet_name, headers)?,
        FileType::Csv => convert_csv(&mut conn, input_path, headers)?,
    }

    Ok(())
}

/// Excel转换，支持指定工作表和自定义表头
fn convert_excel<R>(
    conn: &mut Connection,
    path: &Path,
    sheet_name: Option<&str>,
    headers: Option<&Vec<String>>,
) -> Result<(), Box<dyn std::error::Error>>
where
    R: Reader<std::io::BufReader<std::fs::File>>,
    <R as Reader<std::io::BufReader<std::fs::File>>>::Error: std::error::Error + std::fmt::Display + 'static,
{
    let mut workbook: R = open_workbook(path)?;

    let sheet_names = workbook.sheet_names();
    let sheet_name = if let Some(name) = sheet_name {
        name.to_string()
    } else {
        sheet_names.get(0)
            .ok_or_else(|| "Excel没有工作表".to_string())?
            .to_string()
    };

    let sheet = workbook
        .worksheet_range(&sheet_name)
        .ok_or_else(|| format!("工作表 '{}' 不存在", sheet_name))?;

    let range = sheet?;

    let mut rows = range.rows();

    let header_row = if let Some(custom_headers) = headers {
        custom_headers.clone()
    } else if let Some(row) = rows.next() {
        row.iter().map(|cell| cell.to_string()).collect()
    } else {
        return Err("Excel文件为空".into());
    };

    let table_name = "data";
    create_table(conn, table_name, &header_row)?;

    let placeholders: Vec<String> = (0..header_row.len()).map(|_| "?".to_string()).collect();
    let insert_sql = format!(
        "INSERT INTO {} ({}) VALUES ({})",
        table_name,
        header_row.iter().map(|h| format!("\"{}\"", h)).collect::<Vec<_>>().join(", "),
        placeholders.join(", ")
    );

    let tx = conn.transaction()?;
    {
        let mut stmt = tx.prepare(&insert_sql)?;

        // 如果没自定义表头，第一行是表头，跳过
        if headers.is_none() {
            rows.next();
        }

        for row in rows {
            let values: Vec<String> = row.iter().map(|cell| cell.to_string()).collect();
            stmt.execute(rusqlite::params_from_iter(values))?;
        }
    }
    tx.commit()?;

    Ok(())
}

/// CSV转换，支持自定义表头
fn convert_csv(
    conn: &mut Connection,
    path: &Path,
    headers: Option<&Vec<String>>,
) -> Result<(), Box<dyn std::error::Error>> {
    let file = File::open(path)?;
    let mut reader = csv::Reader::from_reader(file);

    let header_row = if let Some(custom_headers) = headers {
        custom_headers.clone()
    } else {
        reader.headers()?.iter().map(|h| h.to_string()).collect()
    };

    let table_name = "data";
    create_table(conn, table_name, &header_row)?;

    let placeholders: Vec<String> = (0..header_row.len()).map(|_| "?".to_string()).collect();
    let insert_sql = format!(
        "INSERT INTO {} ({}) VALUES ({})",
        table_name,
        header_row.iter().map(|h| format!("\"{}\"", h)).collect::<Vec<_>>().join(", "),
        placeholders.join(", ")
    );

    let tx = conn.transaction()?;
    {
        let mut stmt = tx.prepare(&insert_sql)?;

        // 如果没自定义表头，跳过csv头
        if headers.is_none() {
            reader.records().next();
        }

        for result in reader.records() {
            let record = result?;
            let values: Vec<String> = record.iter().map(|field| field.to_string()).collect();
            stmt.execute(rusqlite::params_from_iter(values))?;
        }
    }
    tx.commit()?;

    Ok(())
}

/// 创建表
fn create_table(conn: &Connection, table_name: &str, headers: &[String]) -> SqliteResult<()> {
    let columns: Vec<String> = headers
        .iter()
        .map(|header| format!("\"{}\" TEXT", header))
        .collect();

    let create_sql = format!(
        "CREATE TABLE IF NOT EXISTS {} ({})",
        table_name,
        columns.join(", ")
    );

    conn.execute(&create_sql, [])?;
    Ok(())
}
