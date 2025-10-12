use egui::Ui;
use crate::ui_step_modules::UiStepModule;

pub struct CsvImportModule {
    selected_file: Option<String>,
}

impl CsvImportModule {
    pub fn new() -> Self {
        Self {
            selected_file: None,
        }
    }
}

impl UiStepModule for CsvImportModule {
    fn get_title(&self) -> String {
        "Import CSV File".to_string()
    }

    fn draw_ui(&mut self, ui: &mut Ui) {
        ui.heading("Select a CSV file to import");
        ui.add_space(10.0);

        // Display selected file if any
        if let Some(path) = &self.selected_file {
            ui.label(format!("Selected file: {}", path));
            ui.add_space(5.0);
        } else {
            ui.label("No file selected");
            ui.add_space(5.0);
        }

        // Button to open file dialog
        if ui.button("Browse...").clicked() {
            if let Some(path) = rfd::FileDialog::new()
                .add_filter("CSV Files", &["csv"])
                .pick_file()
            {
                self.selected_file = Some(path.display().to_string());
            }
        }

        // Clear selection button
        if self.selected_file.is_some() {
            ui.add_space(5.0);
            if ui.button("Clear").clicked() {
                self.selected_file = None;
            }
        }
    }

    fn is_complete(&self) -> bool {
        self.selected_file.is_some()
    }

    fn reset(&mut self) {
        self.selected_file = None;
    }
}