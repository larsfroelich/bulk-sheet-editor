use crate::ui_step_modules::{SharedState, UiStepModule, parse_cell_reference};
use egui::Ui;
use spreadsheet_ods::{read_ods, write_ods};
use std::cell::RefCell;
use std::path::PathBuf;
use std::rc::Rc;

pub struct BulkCreateModule {
    state: Rc<RefCell<SharedState>>,
    save_path: Option<PathBuf>,
    status_message: Option<String>,
    error_message: Option<String>,
}

impl BulkCreateModule {
    pub fn new(state: Rc<RefCell<SharedState>>) -> Self {
        Self {
            state,
            save_path: None,
            status_message: None,
            error_message: None,
        }
    }

    fn validate_inputs(&self) -> Result<(), String> {
        let state = self.state.borrow();
        if state.csv_rows.is_empty() {
            return Err("Import a CSV file before generating sheets.".to_string());
        }
        if state.odf_path.is_none() || state.selected_sheet.is_none() {
            return Err("Select a template worksheet and sheet.".to_string());
        }
        if state
            .cell_mappings
            .iter()
            .all(|mapping| mapping.cell_ref.trim().is_empty())
        {
            return Err("Assign at least one CSV column to a template cell.".to_string());
        }
        Ok(())
    }

    fn generate_and_save(&mut self, path: PathBuf) {
        match self.build_and_write(&path) {
            Ok(sheet_count) => {
                self.status_message = Some(format!(
                    "Created {} sheet(s) using the template.",
                    sheet_count
                ));
                self.error_message = None;
                self.save_path = Some(path.clone());
                self.state.borrow_mut().last_output_path = Some(path);
            }
            Err(err) => {
                self.error_message = Some(err);
            }
        }
    }

    fn build_and_write(&self, output_path: &PathBuf) -> Result<usize, String> {
        let state = self.state.borrow();
        let template_path = state
            .odf_path
            .clone()
            .ok_or_else(|| "Template worksheet missing".to_string())?;
        let template_sheet_name = state
            .selected_sheet
            .clone()
            .ok_or_else(|| "Template sheet missing".to_string())?;
        let mappings = state
            .cell_mappings
            .iter()
            .filter(|mapping| !mapping.cell_ref.trim().is_empty())
            .cloned()
            .collect::<Vec<_>>();
        let rows = state.csv_rows.clone();
        drop(state);

        if rows.is_empty() {
            return Err("CSV file does not contain data rows.".to_string());
        }
        if mappings.is_empty() {
            return Err("No column mappings configured.".to_string());
        }

        let mut workbook = read_ods(&template_path).map_err(|err| err.to_string())?;
        let template_index = workbook
            .sheet_idx(&template_sheet_name)
            .ok_or_else(|| format!("Sheet '{}' not found in template", template_sheet_name))?;
        let template_sheet = workbook.remove_sheet(template_index);

        for (row_index, row) in rows.iter().enumerate() {
            let mut new_sheet = template_sheet.clone();
            let sheet_title = format!("{} {}", template_sheet_name, row_index + 1);
            new_sheet.set_name(sheet_title);
            for mapping in &mappings {
                if let Some((row_idx, col_idx)) = parse_cell_reference(&mapping.cell_ref)
                    && let Some(value) = row.get(mapping.column_index)
                {
                    new_sheet.set_value(row_idx, col_idx, value.clone());
                }
            }
            workbook.push_sheet(new_sheet);
        }

        write_ods(&mut workbook, output_path).map_err(|err| err.to_string())?;
        Ok(rows.len())
    }
}

impl UiStepModule for BulkCreateModule {
    fn get_title(&self) -> String {
        "Generate Workbook".to_string()
    }

    fn draw_ui(&mut self, ui: &mut Ui) {
        ui.label("Create sheets for each CSV row and save them as an ODF workbook");

        match self.validate_inputs() {
            Ok(_) => {
                let state = self.state.borrow();
                ui.label(format!("Rows ready for export: {}", state.csv_rows.len()));
                if let Some(path) = &state.odf_path {
                    ui.label(format!("Template file: {}", path.display()));
                }
                if let Some(sheet) = &state.selected_sheet {
                    ui.label(format!("Template sheet: {}", sheet));
                }
            }
            Err(reason) => {
                ui.colored_label(egui::Color32::DARK_RED, reason);
                return;
            }
        }

        ui.add_space(10.0);
        ui.horizontal(|ui| {
            if ui.button("Save as…").clicked()
                && let Some(path) = rfd::FileDialog::new()
                    .add_filter("ODS", &["ods"])
                    .set_file_name("bulk_output.ods")
                    .save_file()
            {
                self.generate_and_save(path);
            }
            if let Some(path) = &self.save_path {
                ui.label(path.display().to_string());
            }
        });

        if let Some(message) = &self.status_message {
            ui.colored_label(egui::Color32::DARK_GREEN, message);
        }
        if let Some(error) = &self.error_message {
            ui.colored_label(egui::Color32::DARK_RED, error);
        }
    }

    fn is_complete(&self) -> bool {
        self.state.borrow().last_output_path.is_some()
    }

    fn reset(&mut self) {
        self.save_path = None;
        self.status_message = None;
        self.error_message = None;
        self.state.borrow_mut().last_output_path = None;
    }
}
