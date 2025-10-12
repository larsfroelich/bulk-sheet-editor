use crate::ui_step_modules::{ColumnPreview, SharedState, UiStepModule};
use egui::{ScrollArea, Ui};
use std::cell::RefCell;
use std::path::PathBuf;
use std::rc::Rc;

type CsvPreviewData = (Vec<String>, Vec<Vec<String>>, Vec<ColumnPreview>);

pub struct CsvImportModule {
    state: Rc<RefCell<SharedState>>,
    load_error: Option<String>,
}

impl CsvImportModule {
    pub fn new(state: Rc<RefCell<SharedState>>) -> Self {
        Self {
            state,
            load_error: None,
        }
    }

    fn open_csv(&mut self, path: PathBuf) {
        let has_headers = self.state.borrow().csv_has_headers;
        match load_csv_preview(&path, has_headers) {
            Ok((headers, rows, preview)) => {
                let mut state = self.state.borrow_mut();
                state.csv_path = Some(path);
                state.csv_headers = headers;
                state.csv_rows = rows;
                state.csv_preview = preview;
                state.ensure_cell_mappings();
                self.load_error = None;
            }
            Err(err) => {
                self.load_error = Some(err);
            }
        }
    }

    fn update_headers(&mut self) {
        let path = self.state.borrow().csv_path.clone();
        if let Some(path) = path {
            self.open_csv(path);
        }
    }
}

impl UiStepModule for CsvImportModule {
    fn get_title(&self) -> String {
        "Import CSV".to_string()
    }

    fn draw_ui(&mut self, ui: &mut Ui) {
        let selected_path = self
            .state
            .borrow()
            .csv_path
            .as_ref()
            .map(|p| p.display().to_string())
            .unwrap_or_else(|| "No file selected".to_string());

        ui.label("Choose a CSV file and preview its columns");
        let has_selection = self.state.borrow().csv_path.is_some();
        ui.horizontal(|ui| {
            ui.label(selected_path);
            if ui.button("Browse…").clicked()
                && let Some(path) = rfd::FileDialog::new()
                    .add_filter("CSV", &["csv"])
                    .pick_file()
            {
                self.open_csv(path);
            }
            if has_selection && ui.button("Clear").clicked() {
                self.state.borrow_mut().reset_csv();
            }
        });

        let mut has_headers = self.state.borrow().csv_has_headers;
        if ui
            .checkbox(&mut has_headers, "Treat first row as headers")
            .changed()
        {
            self.state.borrow_mut().csv_has_headers = has_headers;
            self.update_headers();
        }

        if let Some(err) = &self.load_error {
            ui.colored_label(egui::Color32::DARK_RED, err);
        }

        let state_snapshot = self.state.borrow().csv_preview.clone();
        if state_snapshot.is_empty() {
            ui.label("Load a CSV file to see column previews.");
            return;
        }

        ui.add_space(10.0);
        ui.heading("Column previews");
        ui.add_space(5.0);
        ScrollArea::vertical().max_height(250.0).show(ui, |ui| {
            for column in state_snapshot {
                draw_column_preview(ui, &column);
            }
        });
    }

    fn is_complete(&self) -> bool {
        !self.state.borrow().csv_rows.is_empty()
    }

    fn reset(&mut self) {
        self.load_error = None;
        self.state.borrow_mut().reset_csv();
    }
}

fn draw_column_preview(ui: &mut Ui, preview: &ColumnPreview) {
    ui.group(|ui| {
        ui.label(format!("Column {} ({})", preview.index + 1, preview.header));
        if preview.samples.is_empty() {
            ui.label("No sample values available");
        } else {
            for value in &preview.samples {
                ui.label(format!("• {}", value));
            }
        }
    });
    ui.add_space(6.0);
}

fn load_csv_preview(path: &PathBuf, has_headers: bool) -> Result<CsvPreviewData, String> {
    let mut reader = csv::ReaderBuilder::new()
        .has_headers(has_headers)
        .from_path(path)
        .map_err(|err| err.to_string())?;

    let headers: Vec<String> = if has_headers {
        reader
            .headers()
            .map_err(|err| err.to_string())?
            .iter()
            .enumerate()
            .map(|(idx, value)| {
                if value.is_empty() {
                    format!("Column {}", idx + 1)
                } else {
                    value.to_string()
                }
            })
            .collect()
    } else {
        Vec::new()
    };

    let mut rows: Vec<Vec<String>> = Vec::new();
    for record in reader.records() {
        let record = record.map_err(|err| err.to_string())?;
        rows.push(record.iter().map(|cell| cell.to_string()).collect());
    }

    let column_count = if has_headers {
        headers
            .len()
            .max(rows.iter().map(|row| row.len()).max().unwrap_or(0))
    } else {
        rows.iter().map(|row| row.len()).max().unwrap_or(0)
    };

    let headers = if has_headers {
        headers
    } else {
        (0..column_count)
            .map(|index| format!("Column {}", index + 1))
            .collect()
    };

    let mut previews = Vec::new();
    for index in 0..column_count {
        let header = headers.get(index).cloned().unwrap_or_default();
        let samples = rows
            .iter()
            .take(5)
            .map(|row| row.get(index).cloned().unwrap_or_default())
            .collect();
        previews.push(ColumnPreview {
            index,
            header,
            samples,
        });
    }

    Ok((headers, rows, previews))
}
