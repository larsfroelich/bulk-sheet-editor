use crate::ui_step_modules::{
    SharedState, UiStepModule, column_label_from_index, parse_cell_reference,
};
use calamine::{Data, DataType, Reader, open_workbook_auto};
use egui::{ComboBox, Grid, Ui};
use std::cell::RefCell;
use std::collections::HashMap;
use std::path::PathBuf;
use std::rc::Rc;

pub struct OdfImportModule {
    state: Rc<RefCell<SharedState>>,
    load_error: Option<String>,
    sheet_error: Option<String>,
}

impl OdfImportModule {
    pub fn new(state: Rc<RefCell<SharedState>>) -> Self {
        Self {
            state,
            load_error: None,
            sheet_error: None,
        }
    }

    fn open_template(&mut self, path: PathBuf) {
        match read_sheet_names(&path) {
            Ok(sheet_names) => {
                let selected_sheet = sheet_names.first().cloned();
                let mut cell_values = HashMap::new();
                if let Some(sheet) = &selected_sheet {
                    match read_sheet_cells(&path, sheet) {
                        Ok(map) => {
                            cell_values = map;
                            self.sheet_error = None;
                        }
                        Err(err) => {
                            self.sheet_error = Some(err);
                        }
                    }
                }
                let mut state = self.state.borrow_mut();
                state.odf_path = Some(path);
                state.odf_sheet_names = sheet_names;
                state.selected_sheet = selected_sheet;
                state.template_cell_values = cell_values;
                self.load_error = None;
            }
            Err(err) => {
                self.load_error = Some(err);
            }
        }
    }

    fn reload_selected_sheet(&mut self) {
        let (path, sheet_name) = {
            let state = self.state.borrow();
            match (state.odf_path.clone(), state.selected_sheet.clone()) {
                (Some(path), Some(sheet)) => (path, sheet),
                _ => return,
            }
        };
        match read_sheet_cells(&path, &sheet_name) {
            Ok(map) => {
                self.state.borrow_mut().template_cell_values = map;
                self.sheet_error = None;
            }
            Err(err) => {
                self.sheet_error = Some(err);
            }
        }
    }
}

impl UiStepModule for OdfImportModule {
    fn get_title(&self) -> String {
        "Configure Template".to_string()
    }

    fn draw_ui(&mut self, ui: &mut Ui) {
        let template_path = self
            .state
            .borrow()
            .odf_path
            .as_ref()
            .map(|p| p.display().to_string())
            .unwrap_or_else(|| "No workbook selected".to_string());

        ui.label("Select a workbook (ODS/XLSX) and map CSV columns to template cells");
        let has_template = self.state.borrow().odf_path.is_some();
        ui.horizontal(|ui| {
            ui.label(template_path);
            if ui.button("Browse…").clicked()
                && let Some(path) = rfd::FileDialog::new()
                    .add_filter("Spreadsheets", &["ods", "xlsx", "xlsm", "xls"])
                    .pick_file()
            {
                self.open_template(path);
            }
            if has_template && ui.button("Clear").clicked() {
                self.state.borrow_mut().reset_template();
            }
        });

        if let Some(err) = &self.load_error {
            ui.colored_label(egui::Color32::DARK_RED, err);
        }

        if self.state.borrow().odf_path.is_none() {
            ui.label("Select a workbook file to continue.");
            return;
        }

        let sheet_names = self.state.borrow().odf_sheet_names.clone();
        if sheet_names.is_empty() {
            ui.label("No sheets found in the selected workbook.");
            return;
        }

        let mut selected_sheet = self.state.borrow().selected_sheet.clone();
        let mut needs_reload = false;
        ComboBox::from_label("Template sheet")
            .selected_text(
                selected_sheet
                    .clone()
                    .unwrap_or_else(|| "Select".to_string()),
            )
            .show_ui(ui, |ui| {
                for sheet in &sheet_names {
                    if ui
                        .selectable_value(&mut selected_sheet, Some(sheet.clone()), sheet)
                        .clicked()
                    {
                        let mut state = self.state.borrow_mut();
                        state.selected_sheet = Some(sheet.clone());
                        needs_reload = true;
                    }
                }
            });

        if needs_reload || self.state.borrow().template_cell_values.is_empty() {
            self.reload_selected_sheet();
        }

        if let Some(err) = &self.sheet_error {
            ui.colored_label(egui::Color32::DARK_RED, err);
        }

        if self.state.borrow().csv_headers.is_empty() {
            ui.label("Import a CSV file to configure column mappings.");
            return;
        }

        let mut state = self.state.borrow_mut();
        let headers = state.csv_headers.clone();
        let first_row = state.csv_rows.first().cloned().unwrap_or_default();
        let template_values = state.template_cell_values.clone();
        let mapping_len = state.cell_mappings.len();
        ui.add_space(10.0);
        ui.heading("Column to cell mapping");
        ui.add_space(5.0);
        Grid::new("column_cell_mapping")
            .striped(true)
            .show(ui, |ui| {
                ui.label("CSV column");
                ui.label("Template cell");
                ui.label("Current value");
                ui.label("New value");
                ui.end_row();

                for index in 0..mapping_len {
                    let mapping = &mut state.cell_mappings[index];
                    let header = headers
                        .get(mapping.column_index)
                        .cloned()
                        .unwrap_or_else(|| format!("Column {}", mapping.column_index + 1));
                    ui.label(header);

                    let mut cell_ref = mapping.cell_ref.clone();
                    if ui.text_edit_singleline(&mut cell_ref).changed() {
                        mapping.cell_ref = cell_ref.trim().to_ascii_uppercase();
                    }

                    let is_valid_cell = parse_cell_reference(&mapping.cell_ref).is_some();
                    let existing = if !is_valid_cell {
                        "(invalid cell)".to_string()
                    } else {
                        template_values
                            .get(&mapping.cell_ref)
                            .cloned()
                            .unwrap_or_else(|| "(empty)".to_string())
                    };
                    ui.label(existing);

                    let new_value = first_row
                        .get(mapping.column_index)
                        .cloned()
                        .unwrap_or_default();
                    ui.label(new_value);
                    ui.end_row();
                }
            });
    }

    fn is_complete(&self) -> bool {
        let state = self.state.borrow();
        state.odf_path.is_some()
            && state.selected_sheet.is_some()
            && state
                .cell_mappings
                .iter()
                .all(|mapping| !mapping.cell_ref.trim().is_empty())
    }

    fn reset(&mut self) {
        self.load_error = None;
        self.sheet_error = None;
        self.state.borrow_mut().reset_template();
    }
}

fn read_sheet_names(path: &PathBuf) -> Result<Vec<String>, String> {
    let workbook = open_workbook_auto(path).map_err(|err| err.to_string())?;
    Ok(workbook.sheet_names().to_vec())
}

fn read_sheet_cells(path: &PathBuf, sheet: &str) -> Result<HashMap<String, String>, String> {
    let mut workbook = open_workbook_auto(path).map_err(|err| err.to_string())?;
    let range = workbook
        .worksheet_range(sheet)
        .map_err(|err| err.to_string())?;

    let mut values = HashMap::new();
    for (row, col, value) in range.cells() {
        if value.is_empty() {
            continue;
        }
        let label = format!("{}{}", column_label_from_index(col as u32), row + 1);
        values.insert(label, stringify_data(value));
    }
    Ok(values)
}

fn stringify_data(data: &Data) -> String {
    match data {
        Data::String(value) => value.clone(),
        Data::Float(value) => format!("{}", value),
        Data::Int(value) => value.to_string(),
        Data::Bool(value) => value.to_string(),
        Data::DateTimeIso(value) | Data::DurationIso(value) => value.clone(),
        Data::DateTime(value) => format!("{:?}", value),
        Data::Error(err) => format!("Error: {:?}", err),
        Data::Empty => String::new(),
    }
}
