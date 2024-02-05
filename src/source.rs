use inquire;
use std::error::Error;
use std::fmt;
use std::fs;
use std::iter;

use crate::config::CONFIG_FILE_NAME;

/// 文件选项
#[derive(Debug)]
struct FileOption {
    value: String,
    label: String,
}

impl fmt::Display for FileOption {
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        write!(f, "{}", self.label)
    }
}

/// 从用户目录或交互式选择获取数据源文件的路径
pub fn get_source_file() -> Result<String, Box<dyn Error>> {
    // 尝试找到默认的源文件
    let file_name = String::from("source.xlsx");
    if fs::metadata(&file_name).is_ok() {
        return Ok(file_name);
    }

    // 如果默认文件不存在，通知用户
    println!("未发现默认数据文件 source.xlsx");

    // 选择其他文件的占位选择值
    let value_other = "other";

    // 搜索当前目录的所有.xslx文件，并添加选项以选取其他文件
    let xlsxes = fs::read_dir(".")?
        .filter_map(Result::ok)
        .filter_map(|entry| {
            let name = entry.file_name().into_string().ok()?;
            if name.ends_with(".xlsx") && name != CONFIG_FILE_NAME {
                Some(FileOption {
                    value: name.clone(),
                    label: name,
                })
            } else {
                None
            }
        })
        .chain(iter::once(FileOption {
            value: value_other.to_string(),
            label: String::from("上面没有，选择其它文件"),
        }))
        .collect::<Vec<_>>();

    // 通过交互式菜单让用户选择文件
    let selected = inquire::Select::new("请选择数据文件", xlsxes).prompt()?;
    let mut value = selected.value;

    // 如果用户选择了其他，提供一个文本接口来输入文件路径
    if value == value_other {
        let messages = vec![
            String::from("选择文件"),
            "=".repeat(50),
            String::from("\n"),
            " ".repeat(16),
            String::from("请拖拽一个文件到此处"),
            String::from("\n"),
            "=".repeat(50),
            String::from("\n"),
        ];
        let message = messages.join("\n");
        let text = inquire::Text::new(&message).prompt()?;
        value = text.replace("'", "");
    }

    // 返回用户选择或输入的文件路径
    Ok(value)
}
