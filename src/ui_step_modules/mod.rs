mod bulk_create;
mod csv_import;
mod odf_import;
mod shared_state;

pub use bulk_create::BulkCreateModule;
pub use csv_import::CsvImportModule;
pub use odf_import::OdfImportModule;
pub use shared_state::{ColumnPreview, SharedState, column_label_from_index, parse_cell_reference};

pub trait UiStepModule {
    fn get_title(&self) -> String;
    fn draw_ui(&mut self, ui: &mut egui::Ui);
    fn is_complete(&self) -> bool;
    fn reset(&mut self);
}
