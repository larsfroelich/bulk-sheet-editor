mod csv_import;
mod test_ui_module;

pub use csv_import::CsvImportModule;
pub use test_ui_module::TestUiModule;

pub trait UiStepModule {
    fn get_title(&self) -> String;
    fn draw_ui(&mut self, ui: &mut egui::Ui);
    fn is_complete(&self) -> bool;
    fn reset(&mut self);
}