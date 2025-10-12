mod bulk_creation;
mod csv_import;
mod odf_template;

pub use bulk_creation::BulkCreationModule;
pub use csv_import::CsvImportModule;
pub use odf_template::OdfTemplateModule;

#[derive(Default, Clone)]
pub struct CsvPreview {
    pub headers: Vec<String>,
    pub sample_rows: Vec<Vec<String>>,
}

#[derive(Default, Clone)]
pub struct CellPreview {
    pub address: String,
    pub value: String,
}

#[derive(Default, Clone)]
pub struct ColumnMapping {
    pub column_index: usize,
    pub column_label: String,
    pub cell_address: String,
}

#[derive(Default, Clone)]
pub struct WorkflowState {
    pub csv_preview: Option<CsvPreview>,
    pub template_sheet_name: Option<String>,
    pub template_cells: Vec<CellPreview>,
    pub column_mappings: Vec<ColumnMapping>,
    pub generated_summary: Option<GenerationSummary>,
}

#[derive(Default, Clone)]
pub struct GenerationSummary {
    pub sheet_count: usize,
    pub output_path: String,
}

pub trait UiStepModule {
    fn get_title(&self) -> String;
    fn draw_ui(&mut self, ui: &mut egui::Ui, state: &mut WorkflowState);
    fn is_complete(&self, state: &WorkflowState) -> bool;
    fn reset(&mut self, state: &mut WorkflowState);
}
