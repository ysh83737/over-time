use std::error::Error;
use time::macros::time;
use time::{Time, Duration};
use time::macros::format_description;

use crate::config::TimeRange;

/// 计算剔除用餐时间的加班时长（小时，保留2位小数）
pub fn calc_overtime_exclude_dinner(start: &String, end: &String, dinner_times: &Vec<TimeRange>) -> Result<Duration, Box<dyn Error>> {
  let format = format_description!("[hour]:[minute]:[second]");
  // 起始时间
  let in_start = Time::parse(&start, &format)?;
  // 结束时间
  let in_end = Time::parse(&end, &format)?;

  // 初始时间
  let zero = time!(00:00:00);

  // 起始时间与初始时间的差
  let duration_start = in_start - zero;
  // 结束时间与初始时间的差
  let mut duration_end = in_end - zero;

  if duration_start > duration_end {
    // 跨日，结束时间加24小时计算
    duration_end += Duration::hours(24);
  }

  // 总加班时间
  let mut result = duration_end - duration_start;

  for item in dinner_times {
    // 晚餐开始时间
    let dinner_start = Time::parse(&item.start, format)? - zero;
    // 晚餐结束时间
    let dinner_end = Time::parse(&item.end, format)? - zero;
    // 晚餐时间长度
    let dinner_diff = dinner_end - dinner_start;

    if duration_start < dinner_start {
      if duration_end > dinner_start {
        if duration_end < dinner_end {
          // --------s-------|--------e-------|---------------
          result -= duration_end - dinner_start;
        } else {
          // --------s-------|---------------|--------e-------
          result -= dinner_diff;
        }
      }
    } else if duration_start < dinner_end {
      if duration_end < dinner_end {
        // ---------------|---s---------e---|---------------
        result = Duration::hours(0);
      } else {
        // ---------------|--------s-------|--------e-------
        result -= dinner_end - duration_start;
      }
    }
  }

  return Ok(result);
}