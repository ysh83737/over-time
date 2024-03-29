use std::error::Error;

use crate::config::TimeRange;
use crate::dinner::calc_overtime_exclude_dinner;

pub struct User {
    pub id: String,
    pub name: String,
    /// 部门
    pub depart: String,
    pub normal: f64,
    pub weekend: f64,
    pub holiday: f64,
    pub total: f64,
}

impl User {
    pub fn new(id: String, name: String, depart: String) -> User {
        User {
            id,
            name,
            depart,
            normal: 0.0,
            weekend: 0.0,
            holiday: 0.0,
            total: 0.0,
        }
    }
    pub fn add(
        &mut self,
        otype: &String,
        start: &String,
        end: &String,
        dinner_times: &Vec<TimeRange>,
    ) -> Result<f64, Box<dyn Error>> {
        // 计算实际加班时长，保留2位小数
        let time = (calc_overtime_exclude_dinner(start, end, dinner_times)?.as_seconds_f64()
            / 60.0
            / 60.0
            * 100.0)
            .round()
            / 100.0;

        match otype.as_str() {
            "法定" => {
                self.holiday += time;
            }
            "周末" => {
                self.weekend += time;
            }
            "其他" => {
                self.normal += time;
            }
            _ => {
                //
            }
        }

        self.total += time;

        return Ok(time);
    }
}
