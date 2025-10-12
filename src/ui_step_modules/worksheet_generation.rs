use crate::ui_step_modules::{
    CsvPreviewState, CsvToTemplateMapping, SharedWorkflowState, TemplatePreviewState, UiStepModule,
};
use alloc::{format, rc::Rc, string::String, vec::Vec};
use core::cell::RefCell;
use csv::ReaderBuilder;
use egui::{Color32, RichText, Ui};
use std::fs::File;
use std::io::{BufReader, Write};
use zip::CompressionMethod;
use zip::write::FileOptions;

const MINIMAL_STYLES: &str = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<office:document-styles xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" office:version=\"1.2\"><office:styles/><office:automatic-styles/><office:master-styles/></office:document-styles>";
const MINIMAL_MANIFEST: &str = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<manifest:manifest xmlns:manifest=\"urn:oasis:names:tc:opendocument:xmlns:manifest:1.0\"><manifest:file-entry manifest:media-type=\"application/vnd.oasis.opendocument.spreadsheet\" manifest:full-path=\"/\"/><manifest:file-entry manifest:media-type=\"text/xml\" manifest:full-path=\"content.xml\"/><manifest:file-entry manifest:media-type=\"text/xml\" manifest:full-path=\"styles.xml\"/></manifest:manifest>";

pub struct WorksheetGenerationModule {
    shared_state: Rc<RefCell<SharedWorkflowState>>,
    sheet_prefix: String,
    last_error: Option<String>,
}

impl WorksheetGenerationModule {
    pub fn new(shared_state: Rc<RefCell<SharedWorkflowState>>) -> Self {
        Self {
            shared_state,
            sheet_prefix: "Sheet".to_string(),
            last_error: None,
        }
    }

    fn csv_state(&self) -> Option<CsvPreviewState> {
        self.shared_state.borrow().csv_preview.clone()
    }

    fn template_state(&self) -> Option<TemplatePreviewState> {
        self.shared_state.borrow().template_preview.clone()
    }

    fn load_csv_rows(path: &str, uses_headers: bool) -> Result<Vec<Vec<String>>, String> {
        let file = File::open(path).map_err(|err| format!("Unable to reopen CSV: {err}"))?;
        let mut reader = ReaderBuilder::new()
            .has_headers(false)
            .from_reader(BufReader::new(file));
        let mut rows: Vec<Vec<String>> = Vec::new();
        for record in reader.records() {
            let record = record.map_err(|err| format!("Failed to parse CSV row: {err}"))?;
            rows.push(record.iter().map(|value| value.to_string()).collect());
        }
        if uses_headers && !rows.is_empty() {
            rows.remove(0);
        }
        Ok(rows)
    }

    fn build_sheet_content(
        mappings: &[CsvToTemplateMapping],
        csv_row: &[String],
    ) -> Vec<((usize, usize), String)> {
        let mut cells = Vec::new();
        for mapping in mappings {
            if let (Some(value), Some(position)) = (
                csv_row.get(mapping.column_index),
                parse_cell_address(&mapping.cell_address),
            ) {
                cells.push((position, value.clone()));
            }
        }
        cells
    }

    fn build_content_xml(
        template: &TemplatePreviewState,
        csv_rows: &[Vec<String>],
        sheet_prefix: &str,
    ) -> String {
        let mut xml = String::from("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n");
        xml.push_str("<office:document-content xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" xmlns:style=\"urn:oasis:names:tc:opendocument:xmlns:style:1.0\" xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" xmlns:table=\"urn:oasis:names:tc:opendocument:xmlns:table:1.0\" office:version=\"1.2\">\n");
        xml.push_str("  <office:body>\n    <office:spreadsheet>\n");

        for (index, row) in csv_rows.iter().enumerate() {
            let sheet_name = format!("{} {}", sheet_prefix, index + 1);
            let cells = Self::build_sheet_content(&template.mappings, row);
            xml.push_str(&format!(
                "      <table:table table:name=\"{}\">\n",
                escape_xml(&sheet_name)
            ));

            if cells.is_empty() {
                xml.push_str("        <table:table-row><table:table-cell/></table:table-row>\n");
            } else {
                let max_row = cells.iter().map(|(pos, _)| pos.0).max().unwrap_or(0);
                let max_col = cells.iter().map(|(pos, _)| pos.1).max().unwrap_or(0);
                for row_index in 0..=max_row {
                    xml.push_str("        <table:table-row>\n");
                    for col_index in 0..=max_col {
                        if let Some((_, value)) = cells
                            .iter()
                            .find(|(pos, _)| pos.0 == row_index && pos.1 == col_index)
                        {
                            if value.is_empty() {
                                xml.push_str("          <table:table-cell/>\n");
                            } else {
                                xml.push_str(&format!(
                                    "          <table:table-cell office:value-type=\"string\"><text:p>{}</text:p></table:table-cell>\n",
                                    escape_xml(value)
                                ));
                            }
                        } else {
                            xml.push_str("          <table:table-cell/>\n");
                        }
                    }
                    xml.push_str("        </table:table-row>\n");
                }
            }
            xml.push_str("      </table:table>\n");
        }

        xml.push_str("    </office:spreadsheet>\n  </office:body>\n</office:document-content>\n");
        xml
    }

    fn write_ods(path: &str, content_xml: &str) -> Result<(), String> {
        let file =
            File::create(path).map_err(|err| format!("Unable to create output file: {err}"))?;
        let mut writer = zip::ZipWriter::new(file);
        let stored = FileOptions::default().compression_method(CompressionMethod::Stored);
        writer
            .start_file("mimetype", stored)
            .map_err(|err| format!("Failed to write mimetype: {err}"))?;
        writer
            .write_all(b"application/vnd.oasis.opendocument.spreadsheet")
            .map_err(|err| format!("Failed to write mimetype: {err}"))?;

        let deflated = FileOptions::default().compression_method(CompressionMethod::Deflated);
        writer
            .start_file("content.xml", deflated)
            .map_err(|err| format!("Failed to start content.xml: {err}"))?;
        writer
            .write_all(content_xml.as_bytes())
            .map_err(|err| format!("Failed to write content.xml: {err}"))?;

        writer
            .start_file("styles.xml", deflated)
            .map_err(|err| format!("Failed to start styles.xml: {err}"))?;
        writer
            .write_all(MINIMAL_STYLES.as_bytes())
            .map_err(|err| format!("Failed to write styles.xml: {err}"))?;

        writer
            .add_directory("META-INF/", deflated)
            .map_err(|err| format!("Failed to create META-INF/: {err}"))?;
        writer
            .start_file("META-INF/manifest.xml", deflated)
            .map_err(|err| format!("Failed to start manifest.xml: {err}"))?;
        writer
            .write_all(MINIMAL_MANIFEST.as_bytes())
            .map_err(|err| format!("Failed to write manifest.xml: {err}"))?;

        writer
            .finish()
            .map_err(|err| format!("Failed to finish ODS: {err}"))?;
        Ok(())
    }

    fn update_status(
        &mut self,
        sheets: usize,
        output_path: Option<String>,
        message: Option<String>,
    ) {
        if let Ok(mut state) = self.shared_state.try_borrow_mut() {
            state.worksheet_state.generated_sheet_count = sheets;
            state.worksheet_state.last_output_path = output_path;
            state.worksheet_state.status_message = message;
        }
    }

    fn generate_and_save(&mut self) {
        let Some(csv_state) = self.csv_state() else {
            self.last_error = Some("Import a CSV file first".to_string());
            return;
        };
        let Some(template_state) = self.template_state() else {
            self.last_error = Some("Select a template worksheet first".to_string());
            return;
        };

        let Some(path) = rfd::FileDialog::new()
            .set_file_name("bulk_sheets.ods")
            .add_filter("ODF Worksheet", &["ods"])
            .save_file()
        else {
            return;
        };
        let output_path = path.display().to_string();

        match Self::load_csv_rows(&csv_state.path, csv_state.uses_headers) {
            Ok(csv_rows) => {
                if csv_rows.is_empty() {
                    self.last_error = Some("The CSV does not contain any data rows".to_string());
                    return;
                }

                let content_xml =
                    Self::build_content_xml(&template_state, &csv_rows, &self.sheet_prefix);
                match Self::write_ods(&output_path, &content_xml) {
                    Ok(()) => {
                        self.last_error = None;
                        self.update_status(
                            csv_rows.len(),
                            Some(output_path.clone()),
                            Some("Worksheet generated".to_string()),
                        );
                    }
                    Err(err) => {
                        self.last_error = Some(err.clone());
                        self.update_status(0, None, Some(err));
                    }
                }
            }
            Err(err) => {
                self.last_error = Some(err.clone());
                self.update_status(0, None, Some(err));
            }
        }
    }
}

impl UiStepModule for WorksheetGenerationModule {
    fn get_title(&self) -> String {
        "Generate bulk worksheets".to_string()
    }

    fn draw_ui(&mut self, ui: &mut Ui) {
        let state = self.shared_state.borrow().clone();

        if let Some(error) = &self.last_error {
            ui.colored_label(Color32::from_rgb(173, 46, 46), error);
            ui.add_space(6.0);
        }

        if let Some(csv) = state.csv_preview {
            ui.label(format!(
                "CSV: {} columns, {} rows",
                csv.columns.len(),
                csv.total_rows
                    .saturating_sub(if csv.uses_headers { 1 } else { 0 })
            ));
        } else {
            ui.label(RichText::new("CSV not yet imported").italics());
        }

        if let Some(template) = state.template_preview {
            ui.label(format!(
                "Template: {} (sheet \"{}\")",
                template.path, template.sheet_name
            ));
            ui.label(format!("Configured mappings: {}", template.mappings.len()));
        } else {
            ui.label(RichText::new("Template not selected").italics());
        }

        ui.add_space(12.0);
        ui.horizontal(|ui| {
            ui.label("Generated sheet name prefix");
            if ui.text_edit_singleline(&mut self.sheet_prefix).changed()
                && self.sheet_prefix.trim().is_empty()
            {
                self.sheet_prefix = "Sheet".to_string();
            }
        });

        ui.add_space(12.0);
        if ui.button("Save worksheet as .ods").clicked() {
            self.generate_and_save();
        }

        ui.add_space(10.0);
        let worksheet_state = state.worksheet_state;
        if let Some(status) = worksheet_state.status_message {
            ui.label(status);
        }
        if let Some(path) = worksheet_state.last_output_path {
            ui.label(format!("Last output: {}", path));
        }
        if worksheet_state.generated_sheet_count > 0 {
            ui.label(format!(
                "Generated {} sheet(s) in the previous run",
                worksheet_state.generated_sheet_count
            ));
        }
    }

    fn is_complete(&self) -> bool {
        self.shared_state
            .borrow()
            .worksheet_state
            .last_output_path
            .is_some()
    }

    fn reset(&mut self) {
        self.last_error = None;
        if let Ok(mut state) = self.shared_state.try_borrow_mut() {
            state.worksheet_state = Default::default();
        }
    }
}

fn parse_cell_address(address: &str) -> Option<(usize, usize)> {
    if address.is_empty() {
        return None;
    }

    let mut letters = String::new();
    let mut digits = String::new();
    for ch in address.chars() {
        if ch.is_ascii_alphabetic() {
            letters.push(ch.to_ascii_uppercase());
        } else if ch.is_ascii_digit() {
            digits.push(ch);
        }
    }

    if letters.is_empty() || digits.is_empty() {
        return None;
    }

    let mut column_index = 0usize;
    for ch in letters.chars() {
        column_index *= 26;
        column_index += (ch as u8 - b'A' + 1) as usize;
    }
    column_index = column_index.saturating_sub(1);

    let row_index: usize = digits.parse::<usize>().ok()?.saturating_sub(1);
    Some((row_index, column_index))
}

fn escape_xml(value: &str) -> String {
    value
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}
