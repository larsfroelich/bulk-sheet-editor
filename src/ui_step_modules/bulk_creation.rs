use crate::ui_step_modules::{
    CellPreview, ColumnMapping, GenerationSummary, UiStepModule, WorkflowState,
};
use egui::{Grid, RichText, TextEdit, Ui};
use std::collections::HashMap;
use std::fs::File;
use std::io::Write;
use zip::CompressionMethod;
use zip::write::FileOptions;

pub struct BulkCreationModule {
    output_path: String,
    status_message: Option<String>,
}

impl BulkCreationModule {
    pub fn new() -> Self {
        Self {
            output_path: String::new(),
            status_message: None,
        }
    }

    fn ensure_output_path(&mut self) {
        if self.output_path.is_empty()
            && let Some(path) = rfd::FileDialog::new()
                .set_file_name("bulk_output.ods")
                .save_file()
        {
            self.output_path = path.display().to_string();
        }
    }

    fn save_bulk_document(
        &mut self,
        state: &mut WorkflowState,
    ) -> Result<GenerationSummary, String> {
        let csv_preview = state
            .csv_preview
            .clone()
            .ok_or_else(|| "Import CSV data before generating worksheets.".to_string())?;
        let template_name = state
            .template_sheet_name
            .clone()
            .unwrap_or_else(|| "Template".to_string());

        if state.template_cells.is_empty() {
            return Err("No template cells available.".to_string());
        }
        if state.column_mappings.is_empty() {
            return Err("No column mappings defined.".to_string());
        }

        let sheet_data = build_sheet_data(
            &state.template_cells,
            &state.column_mappings,
            &csv_preview.sample_rows,
        );
        let sheet_names: Vec<String> = (0..sheet_data.len())
            .map(|idx| format!("{} {}", template_name, idx + 1))
            .collect();

        if sheet_names.is_empty() {
            return Err("No sheet data could be generated.".to_string());
        }

        let content_xml = build_content_xml(&sheet_names, &sheet_data);
        let manifest_xml = build_manifest_xml();

        let file = File::create(&self.output_path)
            .map_err(|err| format!("Unable to create file: {}", err))?;
        let mut writer = zip::ZipWriter::new(file);

        // store mandatory mimetype header without compression
        let stored = FileOptions::default().compression_method(CompressionMethod::Stored);
        writer
            .start_file("mimetype", stored)
            .map_err(|err| format!("Unable to write mimetype: {}", err))?;
        writer
            .write_all(b"application/vnd.oasis.opendocument.spreadsheet")
            .map_err(|err| format!("Unable to write mimetype: {}", err))?;

        // write manifest and content with compression
        let deflated = FileOptions::default().compression_method(CompressionMethod::Deflated);
        writer
            .start_file("META-INF/manifest.xml", deflated)
            .map_err(|err| format!("Unable to write manifest: {}", err))?;
        writer
            .write_all(manifest_xml.as_bytes())
            .map_err(|err| format!("Unable to write manifest: {}", err))?;

        writer
            .start_file("content.xml", deflated)
            .map_err(|err| format!("Unable to write content: {}", err))?;
        writer
            .write_all(content_xml.as_bytes())
            .map_err(|err| format!("Unable to write content: {}", err))?;

        writer
            .finish()
            .map_err(|err| format!("Unable to finalise archive: {}", err))?;

        Ok(GenerationSummary {
            sheet_count: sheet_names.len(),
            output_path: self.output_path.clone(),
        })
    }
}

impl UiStepModule for BulkCreationModule {
    fn get_title(&self) -> String {
        "Create bulk worksheets".to_string()
    }

    fn draw_ui(&mut self, ui: &mut Ui, state: &mut WorkflowState) {
        ui.heading("Generate and save the bulk worksheet");
        ui.add_space(10.0);

        if state.csv_preview.is_none() {
            ui.colored_label(
                egui::Color32::DARK_RED,
                "Complete the CSV import before continuing.",
            );
            return;
        }
        if state.template_sheet_name.is_none() {
            ui.colored_label(
                egui::Color32::DARK_RED,
                "Select and configure a template worksheet first.",
            );
            return;
        }
        if state.column_mappings.is_empty() {
            ui.colored_label(
                egui::Color32::DARK_RED,
                "At least one column mapping is required.",
            );
            return;
        }

        let csv_preview = state.csv_preview.clone().unwrap();
        let data_rows = csv_preview.sample_rows.len().max(1);

        ui.label(format!(
            "Generating {} sheet(s) using template '{}'",
            data_rows,
            state
                .template_sheet_name
                .clone()
                .unwrap_or_else(|| "Template".to_string())
        ));

        ui.add_space(10.0);
        ui.label(RichText::new("Configured mappings").strong());
        Grid::new("bulk_mappings_grid")
            .striped(true)
            .show(ui, |ui| {
                ui.label(RichText::new("Cell").underline());
                ui.label(RichText::new("CSV Column").underline());
                ui.end_row();
                for mapping in &state.column_mappings {
                    ui.label(&mapping.cell_address);
                    ui.label(&mapping.column_label);
                    ui.end_row();
                }
            });

        ui.add_space(15.0);
        ui.horizontal(|ui| {
            ui.label("Output file:");
            ui.add(TextEdit::singleline(&mut self.output_path).hint_text("/path/to/output.ods"));
            if ui.button("Browse").clicked() {
                self.ensure_output_path();
            }
        });

        if ui.button("Save worksheet").clicked() {
            if self.output_path.is_empty() {
                self.status_message = Some("Choose an output path first.".to_string());
            } else {
                match self.save_bulk_document(state) {
                    Ok(summary) => {
                        state.generated_summary = Some(summary.clone());
                        self.status_message = Some(format!(
                            "Saved bulk worksheet with {} sheet(s).",
                            summary.sheet_count
                        ));
                    }
                    Err(err) => {
                        self.status_message = Some(err);
                    }
                }
            }
        }

        if let Some(message) = &self.status_message {
            ui.add_space(10.0);
            ui.label(message);
        }

        if let Some(summary) = &state.generated_summary {
            ui.add_space(10.0);
            ui.label(format!(
                "Last generated file: {} ({} sheet(s))",
                summary.output_path, summary.sheet_count
            ));
        }
    }

    fn is_complete(&self, state: &WorkflowState) -> bool {
        state.generated_summary.is_some()
    }

    fn reset(&mut self, state: &mut WorkflowState) {
        self.output_path.clear();
        self.status_message = None;
        state.generated_summary = None;
    }
}

// --- helpers ---------------------------------------------------------------

fn build_sheet_data(
    template_cells: &[CellPreview],
    mappings: &[ColumnMapping],
    csv_rows: &[Vec<String>],
) -> Vec<Vec<Vec<String>>> {
    let mut base_cells: HashMap<(usize, usize), String> = HashMap::new();
    let mut max_row = 0usize;
    let mut max_col = 0usize;

    for cell in template_cells {
        if let Some((row, col)) = cell_to_indices(&cell.address) {
            max_row = max_row.max(row);
            max_col = max_col.max(col);
            base_cells.insert((row, col), cell.value.clone());
        }
    }

    for mapping in mappings {
        if let Some((row, col)) = cell_to_indices(&mapping.cell_address) {
            max_row = max_row.max(row);
            max_col = max_col.max(col);
            base_cells.entry((row, col)).or_default();
        }
    }

    let rows = max_row + 1;
    let cols = max_col + 1;
    let mut sheets = Vec::with_capacity(csv_rows.len().max(1));

    for csv_row in csv_rows {
        let mut sheet = vec![vec![String::new(); cols]; rows];
        for ((row, col), value) in &base_cells {
            sheet[*row][*col] = value.clone();
        }
        for mapping in mappings {
            if let Some((row, col)) = cell_to_indices(&mapping.cell_address) {
                let value = csv_row
                    .get(mapping.column_index)
                    .cloned()
                    .unwrap_or_default();
                sheet[row][col] = value;
            }
        }
        sheets.push(sheet);
    }

    if csv_rows.is_empty() {
        let mut sheet = vec![vec![String::new(); cols]; rows];
        for ((row, col), value) in base_cells {
            sheet[row][col] = value;
        }
        sheets.push(sheet);
    }

    sheets
}

fn build_content_xml(sheet_names: &[String], sheet_data: &[Vec<Vec<String>>]) -> String {
    let mut xml = String::from("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n");
    xml.push_str(
        "<office:document-content \
xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" \
xmlns:table=\"urn:oasis:names:tc:opendocument:xmlns:table:1.0\" \
xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\">\n",
    );
    xml.push_str("  <office:body>\n    <office:spreadsheet>\n");

    for (name, sheet) in sheet_names.iter().zip(sheet_data.iter()) {
        xml.push_str(&format!(
            "      <table:table table:name=\"{}\">\n",
            escape_xml(name)
        ));
        for row in sheet {
            xml.push_str("        <table:table-row>\n");
            for value in row {
                if value.is_empty() {
                    xml.push_str("          <table:table-cell/>\n");
                } else {
                    xml.push_str(
                        "          <table:table-cell office:value-type=\"string\"><text:p>",
                    );
                    xml.push_str(&escape_xml(value));
                    xml.push_str("</text:p></table:table-cell>\n");
                }
            }
            xml.push_str("        </table:table-row>\n");
        }
        xml.push_str("      </table:table>\n");
    }

    xml.push_str("    </office:spreadsheet>\n  </office:body>\n</office:document-content>\n");
    xml
}

fn build_manifest_xml() -> String {
    "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<manifest:manifest \
xmlns:manifest=\"urn:oasis:names:tc:opendocument:xmlns:manifest:1.0\">\n  <manifest:file-entry manifest:full-path=\"/\" manifest:media-type=\"application/vnd.oasis.opendocument.spreadsheet\"/>\n  <manifest:file-entry manifest:full-path=\"content.xml\" manifest:media-type=\"text/xml\"/>\n</manifest:manifest>\n".to_string()
}

fn cell_to_indices(cell: &str) -> Option<(usize, usize)> {
    let mut letters = String::new();
    let mut digits = String::new();
    for ch in cell.chars() {
        if ch.is_ascii_alphabetic() {
            letters.push(ch.to_ascii_uppercase());
        } else if ch.is_ascii_digit() {
            digits.push(ch);
        }
    }
    if letters.is_empty() || digits.is_empty() {
        return None;
    }

    let mut col = 0usize;
    for ch in letters.chars() {
        col = col * 26 + (ch as usize - 'A' as usize + 1);
    }
    let row = digits.parse::<usize>().ok()?;
    Some((row.saturating_sub(1), col.saturating_sub(1)))
}

fn escape_xml(value: &str) -> String {
    value
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
}
