use std::path::PathBuf;

use eframe::egui::{self, CentralPanel, ViewportBuilder};
use eyre::{OptionExt, Result};
use native_dialog::{FileDialog, MessageDialog};

mod parser;

fn main() -> eframe::Result {
    let options = eframe::NativeOptions {
        viewport: ViewportBuilder::default()
            .with_inner_size([240.0, 240.0])
            .with_resizable(false)
            .with_drag_and_drop(true),
        ..Default::default()
    };

    eframe::run_native(
        "Timesheet Manipulator",
        options,
        Box::new(|_cc| Ok(Box::<App>::default())),
    )
}

#[derive(Default)]
struct App;

impl eframe::App for App {
    fn update(&mut self, ctx: &egui::Context, _frame: &mut eframe::Frame) {
        CentralPanel::default().show(ctx, |ui| {
            ui.label("Browse to the data file.");
            ui.label("Once parsing is complete, choose the location to save the new file.");

            if let Err(e) = browse_button(ui) {
                MessageDialog::new()
                    .set_type(native_dialog::MessageType::Error)
                    .set_title("Encountered an error")
                    .set_text(&e.to_string())
                    .show_alert()
                    .expect("Failed to display dialog");
            }
        });
    }
}

fn browse_button(ui: &mut egui::Ui) -> Result<()> {
    if ui.button("Browse").clicked() {
        let input = FileDialog::new()
            .add_filter("Excel Spreadsheet", &["xlsx", "xls"])
            .show_open_single_file()?
            .ok_or_eyre("User closed input file picker")?;

        let records = parser::read_records_from_file(input)?;

        let output = FileDialog::new()
            .set_filename("Report.xlsx")
            .add_filter("Excel Spreadsheet", &["xlsx", "xls"])
            .show_save_single_file()?
            .ok_or_eyre("User closed output file picker")?;

        parser::write_records_to_file(output.clone(), records)?;

        MessageDialog::new()
            .set_type(native_dialog::MessageType::Info)
            .set_title("Completed")
            .set_text(&format!("Completed writing to {}", output.display()))
            .show_alert()
            .expect("Failed to create completion dialog");
    }
    Ok(())
}
