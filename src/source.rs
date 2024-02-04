use std::fmt;
use std::fs;
use std::io;
use inquire;

use crate::config::CONFIG_FILE_NAME;

#[derive(Debug)]
struct Option {
  value: String,
  label: String,
}

impl fmt::Display for Option {
  fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
      write!(f, "{}", self.label)
  }
}

pub fn get_source_file() -> Result<String, io::Error> {
  let file_name = "source.xlsx";
  let file_metadata = fs::metadata(file_name);

  if let Ok(_) = file_metadata {
    return Ok(file_name.to_string());
  }

  println!("未发现默认数据文件 source.xlsx");

  let dirs: Vec<_> = fs::read_dir(".")?.collect();

  let mut xlsxes: Vec<Option> = vec![];

  for dir in dirs {
    let name = match dir?.file_name().into_string() {
      Ok(s) => s,
      Err(err) => return Err(io::Error::new(io::ErrorKind::Other, err.to_str().unwrap())),
    };
    if name.ends_with(".xlsx") && !name.eq(CONFIG_FILE_NAME) {
      let value = name;
      let label = value.clone();
      xlsxes.push(Option { value, label });
    }
  }

  let value_other = String::from("other");
  xlsxes.push(Option {
    value: value_other.clone(),
    label: String::from("上面没有，选择其它文件")
  });

  let mut result = match inquire::Select::new("请选择数据文件", xlsxes).prompt() {
    Ok(r) => r.value,
    Err(_) => return Err(io::Error::new(io::ErrorKind::Other, "选择文件错误")),
  };

  if result.eq(&value_other) {
    result = match inquire::Text::new("选择文件\n=====================================================\n\n\n                请拖拽一个文件到此处\n\n\n=====================================================\n").prompt() {
      Ok(r) => r,
      Err(_) => return Err(io::Error::new(io::ErrorKind::Other, "选择文件错误")),
    }.replace("'", "");
  }

  return Ok(result);
}