use std::path::PathBuf;

use eframe::egui::{self, CentralPanel, ViewportBuilder};
use rfd::FileDialog;

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
struct App {
    input_path: Option<PathBuf>,
    output_path: Option<PathBuf>,
}
impl eframe::App for App {
    fn update(&mut self, ctx: &egui::Context, _frame: &mut eframe::Frame) {
        CentralPanel::default().show(ctx, |ui| {
            ui.label("Browse to the data file.");
            ui.label("Once parsing is complete, choose the location to save the new file.");

            if ui.button("Browse").clicked() {
                if let Some(path) = FileDialog::new()
                    .add_filter("Excel Spreadsheet", &["xls", "xlsx"])
                    .pick_file()
                {
                    self.input_path = Some(path);
                }
            }
            if let Some(path) = &self.input_path {
                if let Some(path) = path.file_name() {
                    ui.horizontal(|ui| {
                        ui.label("Picked file:");
                        ui.monospace(path.to_string_lossy());
                    });
                }
            }
        });
    }
}
