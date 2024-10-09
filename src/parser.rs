use std::path::PathBuf;

use calamine::{open_workbook, DataType, Reader, Xlsx};
use eyre::{OptionExt, Result};
use rust_xlsxwriter::{Workbook, XlsxSerialize};
use serde::Serialize;

/// A record containing the information for a single day's work. This is serialized into the output
/// spreadsheet.
#[derive(Serialize, XlsxSerialize, Default)]
#[xlsx(
    header_format = Format::new()
        .set_border_right(FormatBorder::Thin)
        .set_border_bottom(FormatBorder::Thin)
        .set_background_color("C0C0C0")
        .set_align(FormatAlign::Center)
)]
pub struct DailyRecord {
    /// The employee's unique ID
    /// TODO: This could be stored as a u32
    #[serde(rename = "Emp ID")]
    employee_id: f64,
    /// The employee's Namekey, built from their full name.
    #[serde(rename = "Namekey")]
    namekey: String,
    /// The employee's full name
    #[serde(rename = "Full Name")]
    full_name: String,
    /// The employee's type code.
    #[serde(rename = "Emp Type Code")]
    emp_type_code: String,
    /// The building's code
    #[serde(rename = "Bld")]
    building: String,
    /// YOE hrs
    /// TODO: What does this mean?
    #[serde(rename = "YOE Hrs")]
    yoe_hrs: f64,
    /// YOE rate
    /// TODO: What does this mean?
    #[serde(rename = "YOE Rate")]
    yoe_rate: f64,
    /// The date that the hours were worked.
    #[serde(rename = "Date Worked")]
    date_worked: String,
    /// The total number of hours worked, rounded to the nearest quarter hour.
    #[serde(rename = "Total Hours")]
    #[xlsx(num_format = "0.00")]
    total_hours: f64,
    /// The description of work rendered.
    #[serde(rename = "Work Description")]
    work_description: String,
}

/// The number of days in the spreadsheet.
const DAY_COUNT: usize = 15;
/// A list of the indexes of each column.
const SORT_LIST: [usize; DAY_COUNT] = [0, 1, 10, 11, 12, 13, 14, 2, 3, 4, 5, 6, 7, 8, 9];

pub fn read_records_from_file(file: PathBuf) -> Result<Vec<DailyRecord>> {
    let mut daily_records: Vec<DailyRecord> = Vec::new();

    let mut workbook: Xlsx<_> = open_workbook(file)?;
    // The first worksheet is what will be read from.
    let (_name, sheet) = &workbook.worksheets()[0];

    // The header is the first row, skip it.
    for row in sheet.rows().skip(1) {
        assert!(row.len() > 6, "Too few columns");
        let employee_id = row[0]
            .get_float()
            .ok_or_eyre("Cannot find employee ID (column A)")?;
        let namekey = row[1]
            .get_string()
            .ok_or_eyre("Cannot find namekey (Column B)")?;
        let full_name = row[2]
            .get_string()
            .ok_or_eyre("Cannot find full name (Column C")?;
        let emp_type_code = row[3]
            .get_string()
            .ok_or_eyre("Cannot find employee type code (Column D")?;
        let building = row[4]
            .get_string()
            .ok_or_eyre("Cannot find building code (Column E)")?;
        let yoe_hrs = row[5]
            .get_float()
            .ok_or_eyre("Cannot find YOE hrs (Column F)")?;
        let yoe_rate = row[6]
            .get_float()
            .ok_or_eyre("Cannot find YOE rate (Column G)")?;

        for i in SORT_LIST {
            let mins = row[i + 8].get_float();
            let hours = row[i + 8 + DAY_COUNT].get_float();
            let date = row[i + 8 + DAY_COUNT * 2].get_string();
            let description = row[i + 8 + DAY_COUNT * 3].get_string();

            // Let's assume that there's data for a record as long as there's a date somewhere.
            if let Some(date) = date {
                daily_records.push(DailyRecord {
                    employee_id,
                    namekey: namekey.to_owned(),
                    full_name: full_name.to_owned(),
                    emp_type_code: emp_type_code.to_owned(),
                    building: building.to_owned(),
                    yoe_hrs,
                    yoe_rate,
                    total_hours: mins.unwrap_or(0.0) + hours.unwrap_or(0.0),
                    work_description: description.unwrap_or("Empty Description").to_owned(),
                    date_worked: date.to_owned(),
                });
            }
        }
    }

    Ok(daily_records)
}

pub fn write_records_to_file(file: &PathBuf, records: Vec<DailyRecord>) -> Result<()> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    worksheet.set_serialize_headers::<DailyRecord>(0, 0)?;
    worksheet.serialize(&records)?;
    workbook.save(file)?;

    Ok(())
}
