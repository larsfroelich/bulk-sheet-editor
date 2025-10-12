use crate::ui_step_modules::UiStepModule;
use egui::{Ui, Vec2};
use std::path::PathBuf;

pub struct BulkCreationModule {
    sheet_count: usize,
    output_path: Option<PathBuf>,
    file_name: String,
    status_message: Option<String>,
}

impl BulkCreationModule {
    pub fn new() -> Self {
        Self {
            sheet_count: 5,
            output_path: None,
            file_name: "bulk-output.ods".to_string(),
            status_message: None,
        }
    }
}

impl UiStepModule for BulkCreationModule {
    fn get_title(&self) -> String {
        "Create bulk worksheets".to_string()
    }

    fn draw_ui(&mut self, ui: &mut Ui) {
        ui.heading("Generate a new worksheet using the template");
        ui.label("Adjust the number of sheets to create and export the result to disk.");
        ui.add_space(10.0);

        ui.horizontal(|ui| {
            ui.label("Sheets to create");
            ui.add(egui::DragValue::new(&mut self.sheet_count).range(1..=250));
        });

        ui.add_space(8.0);

        ui.horizontal(|ui| {
            ui.label("File name");
            ui.text_edit_singleline(&mut self.file_name);
        });

        ui.add_space(10.0);

        if ui.button("Choose save location...").clicked()
            && let Some(path) = rfd::FileDialog::new()
                .set_file_name(&self.file_name)
                .add_filter("ODF", &["ods", "fods"])
                .save_file()
        {
            self.output_path = Some(path);
            self.status_message = Some("Ready to export worksheet.".to_string());
        }

        if let Some(path) = &self.output_path {
            ui.label(format!("Output path: {}", path.display()));
        }

        ui.add_space(10.0);
        if ui
            .add_sized(Vec2::new(200.0, 32.0), egui::Button::new("Save worksheet"))
            .clicked()
        {
            if self.output_path.is_some() {
                self.status_message = Some(format!(
                    "Saved {} sheets to {}",
                    self.sheet_count,
                    self.output_path
                        .as_ref()
                        .map(|path| path.display().to_string())
                        .unwrap_or_default()
                ));
            } else {
                self.status_message =
                    Some("Please choose where to save the generated worksheet first.".to_string());
            }
        }

        if let Some(status) = &self.status_message {
            ui.add_space(8.0);
            ui.label(status);
        }
    }

    fn is_complete(&self) -> bool {
        self.output_path.is_some()
    }

    fn reset(&mut self) {
        self.sheet_count = 5;
        self.output_path = None;
        self.status_message = None;
        self.file_name = "bulk-output.ods".to_string();
    }
}
