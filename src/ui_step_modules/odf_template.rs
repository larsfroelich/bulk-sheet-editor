use crate::ui_step_modules::UiStepModule;
use egui::{ComboBox, Ui};
use std::path::PathBuf;

#[derive(Clone)]
struct CellMapping {
    csv_column_index: usize,
    cell_reference: String,
    current_value: String,
    preview_value: String,
}

impl CellMapping {
    fn with_column(csv_column_index: usize, csv_column_label: &str) -> Self {
        Self {
            csv_column_index,
            cell_reference: "A1".to_string(),
            current_value: "(empty)".to_string(),
            preview_value: format!("Sample data from {}", csv_column_label),
        }
    }
}

pub struct OdfTemplateModule {
    selected_file: Option<PathBuf>,
    available_sheets: Vec<String>,
    selected_sheet: Option<usize>,
    csv_columns: Vec<String>,
    mappings: Vec<CellMapping>,
}

impl OdfTemplateModule {
    pub fn new() -> Self {
        Self {
            selected_file: None,
            available_sheets: Vec::new(),
            selected_sheet: None,
            csv_columns: vec![
                "Column 1".to_string(),
                "Column 2".to_string(),
                "Column 3".to_string(),
                "Column 4".to_string(),
            ],
            mappings: Vec::new(),
        }
    }

    fn ensure_sheet_list(&mut self) {
        if self.available_sheets.is_empty() {
            self.available_sheets = vec![
                "Template".to_string(),
                "Calculations".to_string(),
                "Summary".to_string(),
            ];
        }
    }

    fn add_mapping(&mut self) {
        if self.csv_columns.is_empty() {
            return;
        }
        let column_index = self
            .mappings
            .last()
            .map(|mapping| mapping.csv_column_index)
            .unwrap_or_default();
        let column_index = column_index.min(self.csv_columns.len() - 1);
        let csv_column_label = &self.csv_columns[column_index];
        self.mappings
            .push(CellMapping::with_column(column_index, csv_column_label));
    }
}

impl UiStepModule for OdfTemplateModule {
    fn get_title(&self) -> String {
        "Configure template sheet".to_string()
    }

    fn draw_ui(&mut self, ui: &mut Ui) {
        ui.heading("Select an ODF worksheet as template");
        ui.add_space(8.0);

        match &self.selected_file {
            Some(path) => {
                ui.label(format!("Selected worksheet: {}", path.display()));
            }
            None => {
                ui.label("No worksheet selected");
            }
        }

        if ui.button("Browse worksheet...").clicked()
            && let Some(path) = rfd::FileDialog::new()
                .add_filter("ODF", &["ods", "fods"])
                .pick_file()
        {
            self.selected_file = Some(path);
            self.ensure_sheet_list();
        }

        if self.selected_file.is_some() {
            ui.add_space(10.0);
            ui.separator();
            ui.add_space(10.0);

            ui.heading("Template sheet");
            self.ensure_sheet_list();

            ComboBox::from_label("Select template sheet")
                .selected_text(
                    self.selected_sheet
                        .and_then(|idx| self.available_sheets.get(idx))
                        .cloned()
                        .unwrap_or_else(|| "Choose a sheet".to_string()),
                )
                .show_ui(ui, |ui| {
                    for (idx, sheet) in self.available_sheets.iter().enumerate() {
                        ui.selectable_value(&mut self.selected_sheet, Some(idx), sheet);
                    }
                });

            ui.add_space(15.0);
            ui.heading("Map CSV columns to template cells");
            ui.label("Assign which CSV column should populate a cell in the template.");
            ui.add_space(6.0);

            if ui.button("Add mapping").clicked() {
                self.add_mapping();
            }

            ui.add_space(8.0);

            let mut idx = 0;
            while idx < self.mappings.len() {
                let mut remove_requested = false;
                ui.group(|ui| {
                    let mapping = &mut self.mappings[idx];
                    ui.horizontal(|ui| {
                        ui.label(format!("Mapping {}", idx + 1));
                        if ui.button("Remove").clicked() {
                            remove_requested = true;
                        }
                    });

                    ui.horizontal(|ui| {
                        ui.label("CSV column");
                        ComboBox::from_id_salt(("csv_col", idx))
                            .selected_text(
                                self.csv_columns
                                    .get(mapping.csv_column_index)
                                    .cloned()
                                    .unwrap_or_else(|| "Unknown".to_string()),
                            )
                            .show_ui(ui, |ui| {
                                for (col_idx, label) in self.csv_columns.iter().enumerate() {
                                    let mut value = mapping.csv_column_index;
                                    if ui.selectable_value(&mut value, col_idx, label).clicked() {
                                        mapping.csv_column_index = col_idx;
                                        mapping.preview_value =
                                            format!("Sample data from {}", label);
                                    }
                                }
                            });
                    });

                    ui.horizontal(|ui| {
                        ui.label("Template cell");
                        let response = ui.text_edit_singleline(&mut mapping.cell_reference);
                        if response.changed() {
                            mapping.current_value =
                                format!("Existing value in {}", mapping.cell_reference);
                        }
                    });

                    ui.add_space(6.0);
                    ui.label("Preview");
                    ui.horizontal(|ui| {
                        ui.monospace(format!("Current: {}", mapping.current_value));
                        ui.monospace(format!("Replacement: {}", mapping.preview_value));
                    });
                });

                if remove_requested {
                    self.mappings.remove(idx);
                } else {
                    idx += 1;
                }

                ui.add_space(6.0);
            }

            if self.mappings.is_empty() {
                ui.label("No mappings defined yet. Add at least one mapping to proceed.");
            }
        }
    }

    fn is_complete(&self) -> bool {
        self.selected_file.is_some() && self.selected_sheet.is_some() && !self.mappings.is_empty()
    }

    fn reset(&mut self) {
        self.selected_file = None;
        self.available_sheets.clear();
        self.selected_sheet = None;
        self.mappings.clear();
    }
}
