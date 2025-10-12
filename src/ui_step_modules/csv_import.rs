use crate::ui_step_modules::{CsvPreview, UiStepModule, WorkflowState};
use egui::{Grid, RichText, Ui};
use std::fs::File;
use std::io::BufReader;

const MAX_PREVIEW_ROWS: usize = 8;

pub struct CsvImportModule {
    selected_file: Option<String>,
    raw_preview_rows: Vec<Vec<String>>,
    has_headers: bool,
    error_message: Option<String>,
}

impl CsvImportModule {
    pub fn new() -> Self {
        Self {
            selected_file: None,
            raw_preview_rows: Vec::new(),
            has_headers: true,
            error_message: None,
        }
    }

    fn load_preview(&mut self, state: &mut WorkflowState) {
        self.raw_preview_rows.clear();
        self.error_message = None;
        state.csv_preview = None;

        let Some(path) = &self.selected_file else {
            return;
        };

        match File::open(path) {
            Ok(file) => {
                let mut reader = csv::ReaderBuilder::new()
                    .has_headers(false)
                    .from_reader(BufReader::new(file));
                for result in reader.records().take(MAX_PREVIEW_ROWS) {
                    match result {
                        Ok(record) => {
                            self.raw_preview_rows
                                .push(record.iter().map(|s| s.to_string()).collect());
                        }
                        Err(err) => {
                            self.error_message = Some(format!("Failed to read CSV: {}", err));
                            return;
                        }
                    }
                }
                self.update_state_preview(state);
            }
            Err(err) => {
                self.error_message = Some(format!("Unable to open file: {}", err));
            }
        }
    }

    fn update_state_preview(&self, state: &mut WorkflowState) {
        if self.raw_preview_rows.is_empty() {
            return;
        }

        let column_count = self
            .raw_preview_rows
            .iter()
            .map(|row| row.len())
            .max()
            .unwrap_or(0);

        if column_count == 0 {
            return;
        }

        let mut headers = Vec::with_capacity(column_count);
        if self.has_headers {
            let header_row = &self.raw_preview_rows[0];
            headers.extend((0..column_count).map(|idx| {
                header_row
                    .get(idx)
                    .cloned()
                    .unwrap_or_else(|| format!("Column {}", idx + 1))
            }));
        } else {
            headers.extend((0..column_count).map(|idx| format!("Column {}", idx + 1)));
        }

        let mut sample_rows = Vec::new();
        let start_index = if self.has_headers { 1 } else { 0 };
        for row in self
            .raw_preview_rows
            .iter()
            .skip(start_index)
            .take(MAX_PREVIEW_ROWS.saturating_sub(1))
        {
            let mut padded_row = vec![String::new(); column_count];
            for (idx, cell) in row.iter().enumerate() {
                if idx < column_count {
                    padded_row[idx] = cell.clone();
                }
            }
            sample_rows.push(padded_row);
        }

        state.csv_preview = Some(CsvPreview {
            headers,
            sample_rows,
        });
    }
}

impl UiStepModule for CsvImportModule {
    fn get_title(&self) -> String {
        "Import CSV File".to_string()
    }

    fn draw_ui(&mut self, ui: &mut Ui, state: &mut WorkflowState) {
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
        if ui.button("Browse...").clicked()
            && let Some(path) = rfd::FileDialog::new()
                .add_filter("CSV Files", &["csv"])
                .pick_file()
        {
            self.selected_file = Some(path.display().to_string());
            self.load_preview(state);
        }

        // Clear selection button
        if self.selected_file.is_some() {
            ui.add_space(5.0);
            if ui.button("Clear").clicked() {
                self.selected_file = None;
                self.raw_preview_rows.clear();
                self.error_message = None;
                state.csv_preview = None;
            }
        }

        if let Some(err) = &self.error_message {
            ui.colored_label(egui::Color32::DARK_RED, err);
        }

        if !self.raw_preview_rows.is_empty() {
            ui.separator();
            ui.add_space(5.0);
            ui.checkbox(&mut self.has_headers, "Treat first row as headers");
            if ui.button("Refresh preview").clicked() {
                self.update_state_preview(state);
            }
            ui.add_space(10.0);

            if let Some(preview) = &state.csv_preview {
                ui.label(RichText::new("Column Preview").strong());
                ui.add_space(5.0);
                Grid::new("csv_preview_grid").striped(true).show(ui, |ui| {
                    ui.label(RichText::new("Column").underline());
                    ui.label(RichText::new("Example values").underline());
                    ui.end_row();
                    for (idx, header) in preview.headers.iter().enumerate() {
                        ui.label(header);
                        let mut examples = Vec::new();
                        for row in &preview.sample_rows {
                            if let Some(value) = row.get(idx)
                                && !value.is_empty()
                            {
                                examples.push(value.clone());
                            }
                        }
                        if examples.is_empty() {
                            ui.label("(no data)");
                        } else {
                            ui.label(examples.join(", "));
                        }
                        ui.end_row();
                    }
                });
            }
        }
    }

    fn is_complete(&self, state: &WorkflowState) -> bool {
        self.selected_file.is_some() && state.csv_preview.is_some()
    }

    fn reset(&mut self, state: &mut WorkflowState) {
        self.selected_file = None;
        self.raw_preview_rows.clear();
        self.error_message = None;
        state.csv_preview = None;
    }
}
