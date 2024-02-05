use calamine::{open_workbook, DataType, Reader, Xlsx};
use inquire;
use rust_xlsxwriter::{Workbook as WriterWorkbook, XlsxError as WriterError};
use std::error::Error;
use std::fs;
use std::io;

pub const CONFIG_FILE_NAME: &str = "config.xlsx";

/// 时间区间
#[derive(Debug)]
pub struct TimeRange {
    /// 开始时间
    pub start: String,
    /// 结束时间
    pub end: String,
}

/// 加载配置
///
/// 从`config.xlsx`中加载时间范围配置
pub fn load_config() -> Result<Vec<TimeRange>, Box<dyn Error>> {
    let mut workbook = get_config_workbook()?;
    let name = workbook.sheet_names().get(0).cloned().unwrap_or_default();
    let sheet = workbook.worksheet_range(&name)?;

    let dinner_times: Vec<TimeRange> = sheet.rows()
      .collect::<Vec<_>>()
      // 以2行遍历
      .chunks_exact(2)
      .filter_map(|row_chunks|
          // 分别读取B列的值作为时间区间起止
          if let (DataType::String(start), DataType::String(end)) = (&row_chunks[0][1], &row_chunks[1][1]) {
              Some(TimeRange { start: start.to_string(), end: end.to_string() })
          } else {
              None
          })
      .collect();

    Ok(dinner_times)
}

/// 获取配置工作簿
///
/// 打开或在不存在时创建`config.xlsx`
fn get_config_workbook() -> Result<Xlsx<io::BufReader<fs::File>>, Box<dyn Error>> {
    if fs::metadata(CONFIG_FILE_NAME).is_err() {
        println!("未检测到配置文件 config.xlsx");

        let result = inquire::Confirm::new("是否生成默认配置文件？")
            .with_default(true)
            .prompt()
            .unwrap_or(true);

        if result {
            println!("生成中...");
            create_config_workbook()?;
            println!("生成成功！");
            println!("请修改配置后重新执行");
            return Err(Box::new(WriterError::CustomError(String::from(
                "请修改配置后重新执行",
            ))));
        } else {
            return Err(Box::new(WriterError::CustomError(String::from(
                "无配置文件",
            ))));
        }
    }

    let wb: Xlsx<_> = open_workbook(CONFIG_FILE_NAME)?;
    Ok(wb)
}

/// 创建配置工作簿
///
/// 创建`config.xlsx`，包含默认的午餐和晚餐时间
fn create_config_workbook() -> Result<(), WriterError> {
    let datas = vec![
        ("午餐时间（始）", "12:00:00"),
        ("午餐时间（终）", "13:00:00"),
        ("晚餐时间（始）", "18:00:00"),
        ("晚餐时间（终）", "19:00:00"),
    ];

    let mut workbook = WriterWorkbook::new();
    let worksheet = workbook.add_worksheet();

    for (index, (text_a, text_b)) in datas.iter().enumerate() {
        worksheet.write(index as u32, 0, text_a.to_string())?; // col A
        worksheet.write(index as u32, 1, text_b.to_string())?; // col B
    }

    workbook.save(CONFIG_FILE_NAME)
}
