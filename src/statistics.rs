use calamine::{open_workbook, Error as ReaderError, Reader, Xlsx};
use colored::*;
use figures_rs;
use rust_xlsxwriter::{Workbook, Worksheet};
use std::collections::HashMap;
use std::error::Error;
use time::macros::format_description;
use time::OffsetDateTime;

use crate::config::TimeRange;
use crate::user::User;

// 分析提供的Excel文件，统计员工加班情况
pub fn stat_file(file_path: String, dinner_times: &Vec<TimeRange>) -> Result<(), Box<dyn Error>> {
    // 打开Excel工作簿
    let mut workbook: Xlsx<_> = open_workbook(file_path)?;

    // 获取第一个工作表中的所有行
    let worksheet = workbook
        .worksheet_range_at(0)
        .ok_or(ReaderError::from("读取工作表失败"))??;

    let rows: Vec<_> = worksheet.rows().collect();
    let header = rows[0];

    // 初始化列索引
    let (
        mut index_depart,
        mut index_id,
        mut index_name,
        mut index_start,
        mut index_end,
        mut index_otype,
    ) = (0, 0, 0, 0, 0, 0);

    // 创建加班计算工作表
    let mut calc_sheet = Worksheet::new();
    calc_sheet.set_name("加班计算")?;

    // 解析表头，定位各字段所在列
    for (index, cell) in header.iter().enumerate() {
        match cell.get_string() {
            Some("所属部门") => index_depart = index,
            Some("员工Id") => index_id = index,
            Some("姓名") => index_name = index,
            Some("加班开始时间") => index_start = index,
            Some("加班结束时间") => index_end = index,
            Some("加班类型") => index_otype = index,
            _ => {}
        };
        calc_sheet.write_string(0, index as u16, cell.to_string())?;
    }
    calc_sheet.write_string(0, header.len() as u16, "实际加班时长")?;

    // 存储员工记录
    let mut user_records: HashMap<String, User> = HashMap::new();

    // 遍历所有行，计算加班时间
    for (r_index, row) in rows[1..].iter().enumerate() {
        let depart = row[index_depart].to_string();
        let id = row[index_id].to_string();
        let name = row[index_name].to_string();
        let start = row[index_start].to_string();
        let end = row[index_end].to_string();
        let otype = row[index_otype].to_string();

        // 若员工不存在，则初始化并添加到记录中
        let user = user_records
            .entry(id.clone())
            .or_insert_with(|| User::new(id, name, depart));

        // 将员工记录写入加班计算工作表
        for (c_index, cell) in row.iter().enumerate() {
            calc_sheet.write_string((r_index + 1) as u32, c_index as u16, cell.to_string())?;
        }

        // 计算并写入实际加班时长
        let result = user.add(&otype, &start, &end, dinner_times)?;
        calc_sheet.write((r_index + 1) as u32, row.len() as u16, result)?;
    }

    // 创建加班统计工作表
    let mut stat_sheet = Worksheet::new();
    stat_sheet.set_name("加班统计")?;
    // 写入表头
    stat_sheet.write_row(
        0,
        0,
        [
            "员工Id",
            "部门",
            "姓名",
            "节假日加班",
            "周末加班",
            "其他加班",
            "小计",
        ],
    )?;

    // 遍历记录，汇总每个员工的加班情况
    for (r_index, user) in user_records.values().enumerate() {
        let User {
            id,
            depart,
            name,
            normal,
            weekend,
            holiday,
            total,
        } = user;
        let total = (total * 100.0).round() / 100.0;
        // 写入个人信息
        stat_sheet.write_row((r_index + 1) as u32, 0, [id, depart, name])?;
        // 写入统计数据
        stat_sheet.write_row(
            (r_index + 1) as u32,
            3,
            [holiday.clone(), weekend.clone(), normal.clone(), total],
        )?;
    }

    // 获取当前本地时间
    let now = OffsetDateTime::now_local()?;

    // 格式化日期时间，作为文件名的一部分
    let template = format_description!("[year]_[month]_[day]_[hour]_[minute]_[second]");
    let date_time = now.format(template)?;
    let filename = format!("汇总结果_{}.xlsx", date_time);

    // 保存工作簿
    let mut workbook = Workbook::new();
    workbook.push_worksheet(calc_sheet);
    workbook.push_worksheet(stat_sheet);
    workbook.save(&filename)?;

    println!("{} {}", figures_rs::TICK.green(), "处理完成！".green());
    println!("{} 结果导出为：{}", figures_rs::INFO.blue(), filename.blue());

    Ok(())
}
