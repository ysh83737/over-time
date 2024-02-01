use rust_xlsxwriter::XlsxError;

mod config;

fn main() -> Result<(), XlsxError> {
    let dinner_time = config::load_config()?;
    Ok(())
}
