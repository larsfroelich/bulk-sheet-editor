mod bulk_creation;
mod csv_import;
mod odf_template;
pub use bulk_creation::BulkCreationModule;
pub use csv_import::CsvImportModule;
pub use odf_template::OdfTemplateModule;

pub trait UiStepModule {
    fn get_title(&self) -> String;
    fn draw_ui(&mut self, ui: &mut egui::Ui);
    fn is_complete(&self) -> bool;
    fn reset(&mut self);
}
