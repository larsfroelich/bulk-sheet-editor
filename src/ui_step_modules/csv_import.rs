use crate::ui_step_modules::UiStepModule;
use egui::{self, Ui};
use egui_extras::{Column, TableBuilder};
use std::fs::File;
use std::io::BufReader;
use std::path::PathBuf;

const MAX_PREVIEW_ROWS: usize = 6;

#[derive(Default, Clone)]
struct CsvPreview {
    rows: Vec<Vec<String>>,
}

pub struct CsvImportModule {
    selected_file: Option<PathBuf>,
    treat_first_row_as_headers: bool,
    preview: CsvPreview,
    preview_error: Option<String>,
}

impl CsvImportModule {
    pub fn new() -> Self {
        Self {
            selected_file: None,
            treat_first_row_as_headers: true,
            preview: CsvPreview::default(),
            preview_error: None,
        }
    }

    fn load_preview(&mut self) {
        self.preview = CsvPreview::default();
        self.preview_error = None;

        let Some(path) = &self.selected_file else {
            return;
        };

        match File::open(path) {
            Ok(file) => {
                let reader = BufReader::new(file);
                let mut csv_reader = csv::ReaderBuilder::new()
                    .has_headers(false)
                    .from_reader(reader);

                for result in csv_reader.records().take(MAX_PREVIEW_ROWS) {
                    match result {
                        Ok(record) => self
                            .preview
                            .rows
                            .push(record.iter().map(|value| value.to_string()).collect()),
                        Err(err) => {
                            self.preview_error =
                                Some(format!("Failed to parse CSV preview: {}", err));
                            break;
                        }
                    }
                }
            }
            Err(err) => {
                self.preview_error = Some(format!("Failed to open file: {}", err));
            }
        }
    }

    fn column_count(&self) -> usize {
        self.preview
            .rows
            .iter()
            .map(|row| row.len())
            .max()
            .unwrap_or(0)
    }

    fn headers(&self) -> Vec<String> {
        let column_count = self.column_count();
        if self.treat_first_row_as_headers && !self.preview.rows.is_empty() {
            let first_row = &self.preview.rows[0];
            (0..column_count)
                .map(|idx| first_row.get(idx).cloned().unwrap_or_default())
                .collect()
        } else {
            (0..column_count)
                .map(|idx| format!("Column {}", idx + 1))
                .collect()
        }
    }

    fn preview_rows(&self) -> &[Vec<String>] {
        if self.treat_first_row_as_headers && !self.preview.rows.is_empty() {
            &self.preview.rows[1..]
        } else {
            &self.preview.rows
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

        match &self.selected_file {
            Some(path) => {
                ui.label(format!("Selected file: {}", path.display()));
                ui.add_space(5.0);
            }
            None => {
                ui.label("No file selected");
                ui.add_space(5.0);
            }
        }

        if ui.button("Browse...").clicked()
            && let Some(path) = rfd::FileDialog::new()
                .add_filter("CSV Files", &["csv"])
                .pick_file()
        {
            self.selected_file = Some(path.clone());
            self.load_preview();
        }

        if self.selected_file.is_some() {
            ui.add_space(5.0);
            if ui.button("Clear").clicked() {
                self.selected_file = None;
                self.preview = CsvPreview::default();
                self.preview_error = None;
            }
        }

        if self.selected_file.is_some() {
            ui.separator();
            if ui
                .checkbox(
                    &mut self.treat_first_row_as_headers,
                    "Treat first row as headers",
                )
                .clicked()
            {
                // nothing else to do, derived data recalculates on the fly
            }

            if let Some(err) = &self.preview_error {
                ui.colored_label(egui::Color32::DARK_RED, err);
                return;
            }

            if self.preview.rows.is_empty() {
                ui.label("No preview data available. The file might be empty.");
                return;
            }

            ui.add_space(10.0);
            ui.heading("Column preview");

            let headers = self.headers();
            let preview_rows = self.preview_rows();

            TableBuilder::new(ui)
                .striped(true)
                .column(Column::auto())
                .column(Column::auto())
                .column(Column::remainder())
                .header(20.0, |mut header| {
                    header.col(|ui| {
                        ui.strong("Column #");
                    });
                    header.col(|ui| {
                        ui.strong("Header / Label");
                    });
                    header.col(|ui| {
                        ui.strong("Example values");
                    });
                })
                .body(|mut body| {
                    for (index, header_label) in headers.iter().enumerate() {
                        body.row(24.0, |mut row| {
                            row.col(|ui| {
                                ui.label(format!("{}", index + 1));
                            });
                            row.col(|ui| {
                                ui.label(header_label);
                            });
                            row.col(|ui| {
                                let mut examples = preview_rows
                                    .iter()
                                    .filter_map(|row| row.get(index))
                                    .take(3)
                                    .cloned()
                                    .collect::<Vec<_>>();
                                if examples.is_empty() {
                                    ui.label("(no values in preview)");
                                } else {
                                    if preview_rows.len() > examples.len() {
                                        examples.push("…".to_string());
                                    }
                                    ui.label(examples.join(", "));
                                }
                            });
                        });
                    }
                });
        }
    }

    fn is_complete(&self) -> bool {
        self.selected_file.is_some()
    }

    fn reset(&mut self) {
        self.selected_file = None;
        self.preview = CsvPreview::default();
        self.preview_error = None;
        self.treat_first_row_as_headers = true;
    }
}
