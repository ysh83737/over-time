use std::error::Error;

mod config;
mod source;
mod statistics;
mod dinner;
mod user;

fn main() -> Result<(), Box<dyn Error>> {
    let dinner_times = config::load_config()?;

    let source_path = source::get_source_file()?;

    statistics::stat_file(source_path, &dinner_times)?;
    
    Ok(())
}
