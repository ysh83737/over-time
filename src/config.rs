use std::{fs, io::BufReader};
use inquire;
use rust_xlsxwriter::{Workbook, XlsxError};
use calamine::{open_workbook, DataType, Reader, Xlsx};

const CONFIG_FILE_NAME: &str = "config.xlsx";

/// 时间区间
pub struct TimeRange {
  /// 开始时间
  pub start: String,
  /// 结束时间
  pub end: String,
}

pub fn load_config() -> Result<Vec<TimeRange>, XlsxError> {
  let mut workbook = get_config_workbook()?;
  let name = &workbook.sheet_names()[0];
  let sheet = match workbook.worksheet_range(&name) {
    Ok(sheet) => sheet,
    Err(_) => return Err(XlsxError::CustomError(String::from("读取工作表失败")))
  };

  let rows: Vec<_> = sheet.rows().collect();

  let mut dinner_times: Vec<TimeRange> = vec![];

  for index in 0..2 {
    let start_index = index;
    let end_index = index + 1;
    let start_row = &rows[start_index][1];
    let end_row = &rows[end_index][1];
    println!("index={}, star={}, end={}", index, start_row, end_row);
    if let (DataType::String(start), DataType::String(end)) = (start_row, end_row) {
      dinner_times.push(TimeRange { start: start.to_string(), end: end.to_string() });
    }
  }

  return Ok(dinner_times);
}

fn get_config_workbook() -> Result<Xlsx<BufReader<fs::File>>, XlsxError> {
  let file_metadata = fs::metadata(CONFIG_FILE_NAME);
  if let Err(_) = file_metadata {
      println!("未检测到配置文件 config.xlsx");

      let result = inquire::Confirm::new("是否生成默认配置文件？")
        .with_default(true)
        .prompt()
        .unwrap_or_else(|_| true);

      if result {
        println!("生成中...");
        create_config_workbook()?;
        println!("生成成功！");
        println!("请修改配置后重新执行");
        return Err(XlsxError::ChartError(String::from("新建配置文件")));
      } else {
        return Err(XlsxError::CustomError(String::from("无配置文件")));
      }
  }
  if let Ok(wb) = open_workbook(CONFIG_FILE_NAME) {
    return Ok(wb);
  }
  return Err(XlsxError::CustomError(String::from("读取配置文件失败")));
}

fn create_config_workbook() -> Result<(), XlsxError> {
  let datas = vec![
    ("午餐时间（始）", "12:00:00"),
    ("午餐时间（终）", "13:00:00"),
    ("晚餐时间（始）", "18:00:00"),
    ("晚餐时间（终）", "19:00:00"),
  ];

  let mut workbook = Workbook::new();
  let worksheet = workbook.add_worksheet();

  for (index, (text_a, text_b)) in datas.iter().enumerate() {
    worksheet.write(index as u32, 0, text_a.to_string())?; // col B
    worksheet.write(index as u32, 1, text_b.to_string())?; // col A
  }

  workbook.save(CONFIG_FILE_NAME)?;

  Ok(())
}