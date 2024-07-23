// Prevents additional console window on Windows in release, DO NOT REMOVE!!
#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

use parser::{read_records_from_file, write_records_to_file};
use tauri::api::dialog::blocking::FileDialogBuilder;
mod parser;

struct ParseResult {
    success: bool,
    message: Option<String>,
}

#[tauri::command]
async fn parse_worksheet() -> Result<(), String> {
    let input = FileDialogBuilder::new()
        .set_title("Open the spreadsheet from Skyward")
        .add_filter("Excel Spreadsheet", &["xls", "xlsx"])
        .pick_file()
        .ok_or("User cancelled file prompt".to_owned())?;

    let records = read_records_from_file(input).map_err(|e| e.to_string())?;

    let output = FileDialogBuilder::new()
        .set_title("Select the output location")
        .add_filter("Excel Spreadsheet", &["xls", "xlsx"])
        .set_file_name("Report.xlsx")
        .save_file()
        .ok_or("User cancelled file prompt".to_owned())?;

    write_records_to_file(output, records).map_err(|e| e.to_string())?;

    Ok(())
}

fn main() {
    tauri::Builder::default()
        .invoke_handler(tauri::generate_handler![parse_worksheet])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
