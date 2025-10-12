use crate::ui_step_modules::{CellPreview, ColumnMapping, UiStepModule, WorkflowState};
use egui::{ComboBox, Grid, RichText, TextEdit, Ui};

struct SampleSheet {
    name: String,
    cells: Vec<CellPreview>,
}

pub struct OdfTemplateModule {
    selected_file: Option<String>,
    available_sheets: Vec<SampleSheet>,
    pending_cell: String,
    pending_column: Option<usize>,
    status_message: Option<String>,
}

impl OdfTemplateModule {
    pub fn new() -> Self {
        Self {
            selected_file: None,
            available_sheets: Vec::new(),
            pending_cell: String::new(),
            pending_column: None,
            status_message: None,
        }
    }

    fn load_sample_template(&mut self, state: &mut WorkflowState) {
        // reset shared state when selecting a new template file
        state.template_sheet_name = None;
        state.template_cells.clear();
        state.column_mappings.clear();

        self.available_sheets = vec![
            SampleSheet {
                name: "Template".to_string(),
                cells: vec![
                    CellPreview {
                        address: "A1".to_string(),
                        value: "Project Name".to_string(),
                    },
                    CellPreview {
                        address: "B1".to_string(),
                        value: "Owner".to_string(),
                    },
                    CellPreview {
                        address: "A2".to_string(),
                        value: "Description".to_string(),
                    },
                    CellPreview {
                        address: "B2".to_string(),
                        value: "Pending".to_string(),
                    },
                ],
            },
            SampleSheet {
                name: "Metadata".to_string(),
                cells: vec![
                    CellPreview {
                        address: "A1".to_string(),
                        value: "Created".to_string(),
                    },
                    CellPreview {
                        address: "B1".to_string(),
                        value: "Author".to_string(),
                    },
                    CellPreview {
                        address: "A2".to_string(),
                        value: "Tags".to_string(),
                    },
                ],
            },
        ];

        self.status_message =
            Some("Loaded example sheets. Choose the template sheet below.".to_string());
    }

    fn current_sheet_cells(&self, state: &WorkflowState) -> Vec<CellPreview> {
        if let Some(selected_name) = &state.template_sheet_name {
            for sheet in &self.available_sheets {
                if &sheet.name == selected_name {
                    return sheet.cells.clone();
                }
            }
        }
        Vec::new()
    }

    fn apply_sheet_selection(&mut self, state: &mut WorkflowState, sheet_name: String) {
        let cells = self
            .available_sheets
            .iter()
            .find(|sheet| sheet.name == sheet_name)
            .map(|sheet| sheet.cells.clone())
            .unwrap_or_default();
        state.template_sheet_name = Some(sheet_name);
        state.template_cells = cells;
        state.column_mappings.clear();
        self.status_message =
            Some("Template sheet selected. Set up your column mappings.".to_string());
    }

    fn add_mapping(&mut self, state: &mut WorkflowState) {
        let Some(csv_preview) = &state.csv_preview else {
            self.status_message = Some("Import a CSV file first.".to_string());
            return;
        };

        let Some(column_index) = self.pending_column else {
            self.status_message = Some("Select a CSV column to map.".to_string());
            return;
        };

        let cell_address = self.pending_cell.trim().to_uppercase();
        if cell_address.is_empty() {
            self.status_message = Some("Enter a cell address (e.g. A1).".to_string());
            return;
        }

        if state
            .template_cells
            .iter()
            .all(|cell| cell.address.to_uppercase() != cell_address)
        {
            self.status_message = Some("Cell not present in the template preview.".to_string());
            return;
        }

        if let Some(existing) = state
            .column_mappings
            .iter_mut()
            .find(|mapping| mapping.cell_address == cell_address)
        {
            existing.column_index = column_index;
            existing.column_label = csv_preview.headers[column_index].clone();
            self.status_message = Some("Updated existing mapping.".to_string());
        } else {
            state.column_mappings.push(ColumnMapping {
                column_index,
                column_label: csv_preview.headers[column_index].clone(),
                cell_address: cell_address.clone(),
            });
            self.status_message = Some("Added mapping.".to_string());
        }

        self.pending_cell.clear();
    }
}

impl UiStepModule for OdfTemplateModule {
    fn get_title(&self) -> String {
        "Select Template Worksheet".to_string()
    }

    fn draw_ui(&mut self, ui: &mut Ui, state: &mut WorkflowState) {
        ui.heading("Select an ODF worksheet template");
        ui.add_space(10.0);

        if let Some(path) = &self.selected_file {
            ui.label(format!("Selected template: {}", path));
        } else {
            ui.label("No template selected");
        }

        ui.horizontal(|ui| {
            if ui.button("Browse worksheet...").clicked()
                && let Some(path) = rfd::FileDialog::new()
                    .add_filter("OpenDocument", &["ods", "fods"])
                    .pick_file()
            {
                self.selected_file = Some(path.display().to_string());
                self.load_sample_template(state);
            }
            if self.selected_file.is_some() && ui.button("Clear").clicked() {
                self.selected_file = None;
                self.available_sheets.clear();
                state.template_sheet_name = None;
                state.template_cells.clear();
                state.column_mappings.clear();
            }
        });

        if let Some(message) = &self.status_message {
            ui.label(RichText::new(message).color(egui::Color32::DARK_GREEN));
        }

        if self.available_sheets.is_empty() {
            ui.add_space(10.0);
            ui.label("Select a worksheet file to continue.");
            return;
        }

        ui.separator();
        ui.add_space(10.0);

        let previous_selection = state.template_sheet_name.clone();
        ComboBox::from_label("Template sheet")
            .selected_text(
                previous_selection
                    .clone()
                    .unwrap_or_else(|| "Choose a sheet".to_string()),
            )
            .show_ui(ui, |ui| {
                for sheet in &self.available_sheets {
                    ui.selectable_value(
                        &mut state.template_sheet_name,
                        Some(sheet.name.clone()),
                        sheet.name.clone(),
                    );
                }
            });

        if state.template_sheet_name != previous_selection
            && let Some(name) = state.template_sheet_name.clone()
        {
            self.apply_sheet_selection(state, name);
        }

        if state.template_sheet_name.is_none() {
            ui.label("Select the sheet that should act as the template.");
            return;
        }

        let cells = self.current_sheet_cells(state);
        if cells.is_empty() {
            ui.label("No preview cells available for the selected sheet.");
            return;
        }

        ui.add_space(10.0);
        ui.label(RichText::new("Template cell preview").strong());
        Grid::new("template_cells_grid")
            .striped(true)
            .show(ui, |ui| {
                ui.label(RichText::new("Cell").underline());
                ui.label(RichText::new("Current value").underline());
                ui.end_row();
                for cell in &cells {
                    ui.label(&cell.address);
                    ui.label(&cell.value);
                    ui.end_row();
                }
            });

        ui.add_space(15.0);
        ui.label(RichText::new("Column to cell mapping").strong());

        if state.csv_preview.is_none() {
            ui.colored_label(
                egui::Color32::DARK_RED,
                "Import a CSV file first to configure mappings.",
            );
            return;
        }

        ui.horizontal(|ui| {
            if let Some(preview) = &state.csv_preview {
                ComboBox::from_id_salt("mapping_column_selector")
                    .selected_text(
                        self.pending_column
                            .and_then(|idx| preview.headers.get(idx).cloned())
                            .unwrap_or_else(|| "Select column".to_string()),
                    )
                    .show_ui(ui, |ui| {
                        for (idx, header) in preview.headers.iter().enumerate() {
                            ui.selectable_value(&mut self.pending_column, Some(idx), header);
                        }
                    });
            }
            ui.add(
                TextEdit::singleline(&mut self.pending_cell).hint_text("Cell address (e.g. B2)"),
            );
            if ui.button("Add mapping").clicked() {
                self.add_mapping(state);
            }
        });

        if state.column_mappings.is_empty() {
            ui.label("No mappings configured yet.");
            return;
        }

        ui.add_space(10.0);
        Grid::new("mapping_preview_grid")
            .striped(true)
            .show(ui, |ui| {
                ui.label(RichText::new("Cell").underline());
                ui.label(RichText::new("CSV column").underline());
                ui.label(RichText::new("Existing value").underline());
                ui.label(RichText::new("Preview value").underline());
                ui.label(RichText::new("Actions").underline());
                ui.end_row();

                let sample_values = state
                    .csv_preview
                    .as_ref()
                    .and_then(|preview| preview.sample_rows.first());

                let mut remove_address: Option<String> = None;

                for mapping in &state.column_mappings {
                    let existing_value = state
                        .template_cells
                        .iter()
                        .find(|cell| cell.address.eq_ignore_ascii_case(&mapping.cell_address))
                        .map(|cell| cell.value.clone())
                        .unwrap_or_else(|| "".to_string());

                    let new_value = sample_values
                        .and_then(|row| row.get(mapping.column_index))
                        .cloned()
                        .unwrap_or_default();

                    ui.label(&mapping.cell_address);
                    ui.label(&mapping.column_label);
                    ui.label(existing_value);
                    ui.label(if new_value.is_empty() {
                        "(blank)".to_string()
                    } else {
                        new_value
                    });
                    if ui.button("Remove").clicked() {
                        remove_address = Some(mapping.cell_address.clone());
                    }
                    ui.end_row();
                }

                if let Some(address) = remove_address {
                    state.column_mappings.retain(|m| m.cell_address != address);
                }
            });
    }

    fn is_complete(&self, state: &WorkflowState) -> bool {
        state.template_sheet_name.is_some() && !state.column_mappings.is_empty()
    }

    fn reset(&mut self, state: &mut WorkflowState) {
        self.selected_file = None;
        self.available_sheets.clear();
        self.pending_cell.clear();
        self.pending_column = None;
        self.status_message = None;
        state.template_sheet_name = None;
        state.template_cells.clear();
        state.column_mappings.clear();
    }
}
