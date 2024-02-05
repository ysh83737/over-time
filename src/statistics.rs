use std::collections::HashMap;
use calamine::{open_workbook, Reader, Xlsx, Error};
use rust_xlsxwriter::{Workbook, Worksheet};
use time::OffsetDateTime;
use time::macros::format_description;

use crate::config::TimeRange;
use crate::user::User;

pub fn stat_file(file_path: String, dinner_times: &Vec<TimeRange>) -> Result<(), Error> {
  let mut workbook: Xlsx<_> = open_workbook(file_path)?;

  let worksheet = match workbook.worksheet_range_at(0) {
    Some(ws) => ws?,
    None => return Err(Error::from("读取工作表失败")),
  };

  let rows: Vec<_> = worksheet.rows().collect();
  let header = rows[0];

  let mut index_depart = 0;
  let mut index_id = 0;
  let mut index_name = 0;
  let mut index_start = 0;
  let mut index_end = 0;
  let mut index_otype = 0;

  let mut calc_sheet = Worksheet::new();
  let _ = calc_sheet.set_name("加班计算");

  for (index, cell) in header.iter().enumerate() {
    match cell.get_string() {
      Some("所属部门") => index_depart = index,
      Some("员工Id") => index_id = index,
      Some("姓名") => index_name = index,
      Some("加班开始时间") => index_start = index,
      Some("加班结束时间") => index_end = index,
      Some("加班类型") => index_otype = index,
      _ => {},
    };
    let _ = calc_sheet.write(0, index as u16, cell.to_string());
  };
  let _ = calc_sheet.write(0, header.len() as u16, "实际加班时长");
  
  let mut user_records: HashMap<String, User> = HashMap::new();

  for (r_index, row) in rows[1..].iter().enumerate() {
    let depart = &row[index_depart].to_string();
    let id = &row[index_id].to_string();
    let name = &row[index_name].to_string();
    let start = &row[index_start].to_string();
    let end = &row[index_end].to_string();
    let otype = &row[index_otype].to_string();

    let user = match user_records.get_mut(id) {
      Some(u) => u,
      None => {
        let mut user = User::new();
        user.id = id.to_string();
        user.name = name.to_string();
        user.depart = depart.to_string();

        user_records.insert(id.to_string(), user);
        user_records.get_mut(id).unwrap()
      },
    };

    for (c_index, cell) in row.iter().enumerate() {
      let _ = calc_sheet.write((r_index + 1) as u32, c_index as u16, cell.to_string());
    }

    let result = match user.add(otype, start, end, dinner_times) {
      Ok(r) => r,
      Err(_) => return Err(Error::from("计算实际加班时长错误")),
    };
    let _ = calc_sheet.write((r_index + 1) as u32, row.len() as u16, result);
  }

  let mut stat_sheet = Worksheet::new();
  let _ = stat_sheet.set_name("加班统计");
  let _ = stat_sheet.write_row(0, 0, ["员工Id", "部门", "姓名", "节假日加班", "周末加班", "其他加班", "小计"]);

  for (r_index, user) in user_records.values().enumerate() {
    let User { id, depart, name, normal, weekend, holiday, total } = user;
    let total = (total * 100.0).round() / 100.0;
    let _ = stat_sheet.write_row((r_index + 1) as u32, 0, [id, depart, name]);
    let _ = stat_sheet.write_row(
      (r_index + 1) as u32,
      0,
      [holiday.clone(), weekend.clone(), normal.clone(), total],
    );
  }

  let mut workbook = Workbook::new();

  workbook.push_worksheet(calc_sheet);
  workbook.push_worksheet(stat_sheet);

  let now = OffsetDateTime::now_local().unwrap();

  let template = format_description!("[year]_[month]_[day]_[hour]_[minute]_[second]");
  let date_time = now.format(template).unwrap();
  let filename = format!("汇总结果_{}.xlsx", date_time);

  let _ = workbook.save(&filename);

  println!("处理完成！");
  println!("结果导出为：{}", filename);

  return Ok(())
}
