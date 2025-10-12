use crate::ui_step_modules::{
    CsvPreviewState, CsvToTemplateMapping, SharedWorkflowState, TemplatePreviewRow,
    TemplatePreviewState, UiStepModule,
};
use alloc::{format, rc::Rc, string::String, vec, vec::Vec};
use core::cell::RefCell;
use egui::{Color32, RichText, Ui};

const DEFAULT_SHEETS: &[&str] = &["Template", "Sheet1", "Sheet2"];

pub struct OdfTemplateModule {
    shared_state: Rc<RefCell<SharedWorkflowState>>,
    selected_file: Option<String>,
    available_sheets: Vec<String>,
    selected_sheet_index: Option<usize>,
    column_cell_inputs: Vec<String>,
    last_error: Option<String>,
}

impl OdfTemplateModule {
    pub fn new(shared_state: Rc<RefCell<SharedWorkflowState>>) -> Self {
        Self {
            shared_state,
            selected_file: None,
            available_sheets: Vec::new(),
            selected_sheet_index: None,
            column_cell_inputs: Vec::new(),
            last_error: None,
        }
    }

    fn csv_state(&self) -> Option<CsvPreviewState> {
        self.shared_state.borrow().csv_preview.clone()
    }

    fn sync_columns_with_csv(&mut self) {
        let column_count = self
            .csv_state()
            .map(|preview| preview.columns.len())
            .unwrap_or_default();
        if self.column_cell_inputs.len() < column_count {
            self.column_cell_inputs.extend(vec![
                String::new();
                column_count - self.column_cell_inputs.len()
            ]);
        } else if self.column_cell_inputs.len() > column_count {
            self.column_cell_inputs.truncate(column_count);
        }
    }

    fn update_shared_state(&mut self) {
        let csv_preview = self.csv_state();
        let Some(csv_preview) = csv_preview else {
            if let Ok(mut state) = self.shared_state.try_borrow_mut() {
                state.template_preview = None;
            }
            return;
        };

        let Some(path) = &self.selected_file else {
            if let Ok(mut state) = self.shared_state.try_borrow_mut() {
                state.template_preview = None;
            }
            return;
        };

        let Some(sheet_index) = self.selected_sheet_index else {
            if let Ok(mut state) = self.shared_state.try_borrow_mut() {
                state.template_preview = None;
            }
            return;
        };

        let sheet_name = self
            .available_sheets
            .get(sheet_index)
            .cloned()
            .unwrap_or_else(|| "Sheet1".to_string());

        let mut mappings: Vec<CsvToTemplateMapping> = Vec::new();
        let mut preview_rows: Vec<TemplatePreviewRow> = Vec::new();

        for (column_index, cell_address) in self.column_cell_inputs.iter().enumerate() {
            if cell_address.trim().is_empty() {
                continue;
            }

            let column_label = csv_preview.column_label(column_index);
            let preview_value = csv_preview
                .columns
                .get(column_index)
                .and_then(|column| column.sample_values.first().cloned());

            mappings.push(CsvToTemplateMapping {
                column_index,
                cell_address: cell_address.trim().to_string(),
            });

            preview_rows.push(TemplatePreviewRow {
                column_label,
                cell_address: cell_address.trim().to_string(),
                existing_value: Some("preview unavailable".to_string()),
                preview_value,
            });
        }

        if let Ok(mut state) = self.shared_state.try_borrow_mut() {
            state.template_preview = Some(TemplatePreviewState {
                path: path.clone(),
                sheet_name,
                mappings,
                preview_rows,
            });
        }
    }

    fn ensure_sheet_list(&mut self) {
        if self.available_sheets.is_empty() {
            self.available_sheets = DEFAULT_SHEETS
                .iter()
                .map(|entry| (*entry).to_string())
                .collect();
        }
        if self.selected_sheet_index.is_none() && !self.available_sheets.is_empty() {
            self.selected_sheet_index = Some(0);
        }
    }

    fn draw_mapping_editor(&mut self, ui: &mut Ui) {
        self.sync_columns_with_csv();
        let csv_preview = match self.csv_state() {
            Some(preview) => preview,
            None => {
                ui.label(RichText::new("Import a CSV file first").italics());
                return;
            }
        };

        if self.selected_file.is_none() {
            ui.label(RichText::new("Select an ODF worksheet to continue").italics());
            return;
        }

        if self.selected_sheet_index.is_none() {
            ui.label(RichText::new("Choose a template sheet").italics());
            return;
        }

        let mut mapping_changed = false;
        egui::Grid::new("template_mapping_grid")
            .striped(true)
            .show(ui, |ui| {
                ui.heading("CSV column");
                ui.heading("Template cell");
                ui.heading("Preview");
                ui.end_row();

                for (index, input) in self.column_cell_inputs.iter_mut().enumerate() {
                    let column_label = csv_preview.column_label(index);
                    ui.label(RichText::new(column_label.clone()).strong());

                    if ui.text_edit_singleline(input).changed() {
                        mapping_changed = true;
                    }

                    let preview_value = csv_preview
                        .columns
                        .get(index)
                        .and_then(|column| column.sample_values.first())
                        .map(String::as_str)
                        .unwrap_or("<empty>");
                    ui.label(format!("{} → {}", preview_value, input.trim()));
                    ui.end_row();
                }
            });
        if mapping_changed {
            self.update_shared_state();
        }

        ui.add_space(10.0);
        if let Some(state) = self.shared_state.borrow().template_preview.clone() {
            if state.preview_rows.is_empty() {
                ui.label(RichText::new("No mappings configured yet").italics());
            } else {
                ui.collapsing("Preview", |ui| {
                    for row in state.preview_rows {
                        ui.label(row.to_string());
                    }
                });
            }
        }
    }
}

impl UiStepModule for OdfTemplateModule {
    fn get_title(&self) -> String {
        "Select template worksheet".to_string()
    }

    fn draw_ui(&mut self, ui: &mut Ui) {
        self.ensure_sheet_list();

        if let Some(error) = &self.last_error {
            ui.colored_label(Color32::from_rgb(173, 46, 46), error);
            ui.add_space(8.0);
        }

        ui.horizontal(|ui| {
            let chosen_path = if ui.button("Choose worksheet...").clicked() {
                rfd::FileDialog::new()
                    .add_filter("ODF worksheet", &["ods"])
                    .pick_file()
            } else {
                None
            };
            if let Some(path) = chosen_path {
                self.selected_file = Some(path.display().to_string());
                self.last_error = None;
                self.update_shared_state();
            }

            if let Some(path) = &self.selected_file {
                ui.label(path);
            } else {
                ui.label(RichText::new("No template selected").italics());
            }
        });

        ui.add_space(8.0);

        if !self.available_sheets.is_empty() {
            let sheet_names = self.available_sheets.clone();
            let mut current_index = self.selected_sheet_index.unwrap_or(0);
            let mut selection_changed = false;
            egui::ComboBox::from_label("Template sheet")
                .selected_text(
                    sheet_names
                        .get(current_index)
                        .map(String::as_str)
                        .unwrap_or("Select"),
                )
                .show_ui(ui, |ui| {
                    for (idx, name) in sheet_names.iter().enumerate() {
                        if ui.selectable_value(&mut current_index, idx, name).clicked() {
                            selection_changed = true;
                        }
                    }
                });
            if selection_changed {
                self.selected_sheet_index = Some(current_index);
                self.update_shared_state();
            }
        }

        ui.add_space(12.0);
        self.draw_mapping_editor(ui);
    }

    fn is_complete(&self) -> bool {
        self.shared_state
            .borrow()
            .template_preview
            .as_ref()
            .map(|preview| !preview.mappings.is_empty())
            .is_some()
    }

    fn reset(&mut self) {
        self.selected_file = None;
        self.available_sheets.clear();
        self.selected_sheet_index = None;
        self.column_cell_inputs.clear();
        self.last_error = None;
        if let Ok(mut state) = self.shared_state.try_borrow_mut() {
            state.template_preview = None;
        }
    }
}
