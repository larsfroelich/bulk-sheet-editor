use alloc::{string::String, vec::Vec};
use core::fmt;

#[derive(Clone, Debug)]
pub struct CsvColumnPreview {
    pub index: usize,
    pub header: Option<String>,
    pub sample_values: Vec<String>,
}

#[derive(Clone, Debug)]
pub struct CsvPreviewState {
    pub path: String,
    pub uses_headers: bool,
    pub columns: Vec<CsvColumnPreview>,
    pub total_rows: usize,
}

impl CsvPreviewState {
    pub fn column_label(&self, idx: usize) -> String {
        self.columns
            .get(idx)
            .map(|column| {
                column
                    .header
                    .clone()
                    .unwrap_or_else(|| format!("Column {}", idx + 1))
            })
            .unwrap_or_else(|| format!("Column {}", idx + 1))
    }
}

#[derive(Clone, Debug, Default)]
pub struct CsvToTemplateMapping {
    pub column_index: usize,
    pub cell_address: String,
}

#[derive(Clone, Debug)]
pub struct TemplatePreviewRow {
    pub column_label: String,
    pub cell_address: String,
    pub existing_value: Option<String>,
    pub preview_value: Option<String>,
}

impl fmt::Display for TemplatePreviewRow {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(
            f,
            "{} → {} ({} → {})",
            self.column_label,
            self.cell_address,
            self.existing_value
                .as_deref()
                .filter(|value| !value.is_empty())
                .unwrap_or("empty"),
            self.preview_value
                .as_deref()
                .filter(|value| !value.is_empty())
                .unwrap_or("empty")
        )
    }
}

#[derive(Clone, Debug)]
pub struct TemplatePreviewState {
    pub path: String,
    pub sheet_name: String,
    pub mappings: Vec<CsvToTemplateMapping>,
    pub preview_rows: Vec<TemplatePreviewRow>,
}

#[derive(Clone, Debug, Default)]
pub struct WorksheetGenerationState {
    pub last_output_path: Option<String>,
    pub generated_sheet_count: usize,
    pub status_message: Option<String>,
}

#[derive(Default, Clone, Debug)]
pub struct SharedWorkflowState {
    pub csv_preview: Option<CsvPreviewState>,
    pub template_preview: Option<TemplatePreviewState>,
    pub worksheet_state: WorksheetGenerationState,
}
