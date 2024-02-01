use std::fs;
use inquire;

pub fn load_config() {
  get_config_workbook();
}

fn get_config_workbook() -> Result<String, String> {
  let file_name = "config.xlsx";
  let file_metadata = fs::metadata(file_name);
  if let Err(_) = file_metadata {
      println!("未检测到配置文件 config.xlsx");

      let result = inquire::Confirm::new("是否生成默认配置文件？")
        .with_default(true)
        .prompt()
        .unwrap_or_else(|_| true);

      if result {
        println!("生成中...");
        println!("生成成功！");
        return Ok(String::from("生成成功"))
      } else {
        return Err(String::from("无配置文件"))
      }
  }
  return Ok(file_name.to_string());
}
