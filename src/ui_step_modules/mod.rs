mod csv_import;
mod odf_template;
mod shared_state;
mod worksheet_generation;

pub use csv_import::CsvImportModule;
pub use odf_template::OdfTemplateModule;
pub use shared_state::*;
pub use worksheet_generation::WorksheetGenerationModule;

pub trait UiStepModule {
    fn get_title(&self) -> String;
    fn draw_ui(&mut self, ui: &mut egui::Ui);
    fn is_complete(&self) -> bool;
    fn reset(&mut self);
}
