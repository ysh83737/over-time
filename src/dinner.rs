use time::error::Parse;
use time::macros::time;
use time::{Time, Duration};
use time::macros::format_description;

use crate::config::TimeRange;

pub fn calc_overtime_exclude_dinner(start: &String, end: &String, dinner_times: &Vec<TimeRange>) -> Result<Duration, Parse> {
  let format = format_description!("[hour]:[minute]:[second]");
  let in_start = Time::parse(&start, &format)?;
  let in_end = Time::parse(&end, &format)?;

  let zero = time!(00:00:00);

  let duration_start = in_start - zero;
  let mut duration_end = in_end - zero;

  if duration_start > duration_end {
    duration_end += Duration::hours(24);
  }

  let mut result = duration_end - duration_start;

  for item in dinner_times {
    let dinner_start = Time::parse(&item.start, format)? - zero;
    let dinner_end = Time::parse(&item.end, format)? - zero;
    let dinner_diff = dinner_end - dinner_start;

    if duration_start < dinner_start {
      if duration_end > dinner_start {
        if duration_end < dinner_end {
          // --------s-------|--------e-------|---------------
          result -= duration_end - dinner_start;
        } else {
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
