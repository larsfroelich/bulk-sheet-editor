use crate::ui_step_modules::{
    CsvColumnPreview, CsvPreviewState, SharedWorkflowState, UiStepModule,
};
use alloc::{format, rc::Rc, string::String, vec, vec::Vec};
use core::cell::RefCell;
use csv::ReaderBuilder;
use egui::{Color32, RichText, Ui};
use std::fs::File;
use std::io::BufReader;

const SAMPLE_ROWS: usize = 5;

pub struct CsvImportModule {
    shared_state: Rc<RefCell<SharedWorkflowState>>,
    selected_file: Option<String>,
    treat_first_row_as_headers: bool,
    last_error: Option<String>,
}

impl CsvImportModule {
    pub fn new(shared_state: Rc<RefCell<SharedWorkflowState>>) -> Self {
        Self {
            shared_state,
            selected_file: None,
            treat_first_row_as_headers: true,
            last_error: None,
        }
    }

    fn load_preview(&mut self) {
        let Some(path) = &self.selected_file else {
            return;
        };

        match Self::read_csv_preview(path, self.treat_first_row_as_headers) {
            Ok(preview) => {
                self.last_error = None;
                if let Ok(mut state) = self.shared_state.try_borrow_mut() {
                    state.csv_preview = Some(preview);
                }
            }
            Err(err) => {
                self.last_error = Some(err);
            }
        }
    }

    fn read_csv_preview(
        path: &str,
        treat_first_row_as_headers: bool,
    ) -> Result<CsvPreviewState, String> {
        let file = File::open(path).map_err(|err| format!("Failed to open CSV: {err}"))?;
        let mut reader = ReaderBuilder::new()
            .has_headers(false)
            .from_reader(BufReader::new(file));

        let mut rows: Vec<Vec<String>> = Vec::new();
        let mut total_rows = 0usize;
        for record in reader.records() {
            let record = record.map_err(|err| format!("Failed to parse CSV row: {err}"))?;
            total_rows += 1;
            if rows.len() < SAMPLE_ROWS + 1 {
                rows.push(record.iter().map(|value| value.to_string()).collect());
            }
        }

        if rows.is_empty() {
            return Err("The selected CSV does not contain any data".to_string());
        }

        let mut header_row: Vec<String> = Vec::new();
        let mut data_rows = rows;
        let uses_headers = treat_first_row_as_headers;
        if uses_headers {
            header_row = data_rows
                .first()
                .cloned()
                .unwrap_or_else(|| vec![String::new(); data_rows[0].len()]);
            if !data_rows.is_empty() {
                data_rows.remove(0);
            }
        }

        let column_count = core::cmp::max(
            header_row.len(),
            data_rows.iter().map(|row| row.len()).max().unwrap_or(0),
        );

        let mut columns: Vec<CsvColumnPreview> = Vec::with_capacity(column_count);
        for column_index in 0..column_count {
            let header = header_row
                .get(column_index)
                .cloned()
                .filter(|value| !value.is_empty());
            let mut sample_values = Vec::with_capacity(SAMPLE_ROWS);
            for row in data_rows.iter().take(SAMPLE_ROWS) {
                sample_values.push(row.get(column_index).cloned().unwrap_or_default());
            }
            columns.push(CsvColumnPreview {
                index: column_index,
                header,
                sample_values,
            });
        }

        Ok(CsvPreviewState {
            path: path.to_string(),
            uses_headers,
            columns,
            total_rows,
        })
    }

    fn show_preview(&self, ui: &mut Ui) {
        let state = self.shared_state.borrow();
        if let Some(preview) = &state.csv_preview {
            ui.label(format!(
                "Detected {} column(s) and {} data row(s)",
                preview.columns.len(),
                preview
                    .total_rows
                    .saturating_sub(if preview.uses_headers { 1 } else { 0 })
            ));
            ui.add_space(6.0);

            egui::ScrollArea::horizontal().show(ui, |ui| {
                egui::Grid::new("csv_preview_grid")
                    .striped(true)
                    .show(ui, |ui| {
                        ui.heading("Column");
                        ui.heading("Sample values");
                        ui.end_row();

                        for column in &preview.columns {
                            let name = column
                                .header
                                .clone()
                                .filter(|value| !value.is_empty())
                                .unwrap_or_else(|| format!("Column {}", column.index + 1));
                            ui.label(RichText::new(name).strong());

                            if column.sample_values.is_empty() {
                                ui.label(RichText::new("<no sample data>").italics());
                            } else {
                                let display_values = column
                                    .sample_values
                                    .iter()
                                    .map(|value| if value.is_empty() { "<empty>" } else { value })
                                    .collect::<Vec<_>>()
                                    .join(", ");
                                ui.label(display_values);
                            }
                            ui.end_row();
                        }
                    });
            });
        } else {
            ui.label(RichText::new("No CSV file loaded yet.").italics());
        }
    }
}

impl UiStepModule for CsvImportModule {
    fn get_title(&self) -> String {
        "Import CSV file".to_string()
    }

    fn draw_ui(&mut self, ui: &mut Ui) {
        ui.heading("Select a CSV file to import");
        ui.add_space(8.0);

        if let Some(error) = &self.last_error {
            ui.colored_label(Color32::from_rgb(173, 46, 46), error);
            ui.add_space(6.0);
        }

        ui.horizontal(|ui| {
            let chosen_file = if ui.button("Browse...").clicked() {
                rfd::FileDialog::new()
                    .add_filter("CSV files", &["csv"])
                    .pick_file()
            } else {
                None
            };
            if let Some(path) = chosen_file {
                self.selected_file = Some(path.display().to_string());
                self.load_preview();
            }

            if let Some(path) = &self.selected_file {
                ui.label(path);
            } else {
                ui.label(RichText::new("No file selected").italics());
            }
        });

        if ui
            .checkbox(
                &mut self.treat_first_row_as_headers,
                "Treat first row as headers",
            )
            .changed()
        {
            self.load_preview();
        }

        if ui.button("Reload preview").clicked() {
            self.load_preview();
        }

        ui.add_space(12.0);
        self.show_preview(ui);
    }

    fn is_complete(&self) -> bool {
        self.shared_state
            .borrow()
            .csv_preview
            .as_ref()
            .map(|preview| !preview.columns.is_empty())
            .is_some()
    }

    fn reset(&mut self) {
        self.selected_file = None;
        self.last_error = None;
        if let Ok(mut state) = self.shared_state.try_borrow_mut() {
            state.csv_preview = None;
        }
    }
}
