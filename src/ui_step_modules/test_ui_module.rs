use egui::Ui;
use crate::ui_step_modules::UiStepModule;

pub struct TestUiModule {
    completed: bool
}

impl TestUiModule {
    pub fn new() -> Self {
        Self {
            completed: false,
        }
    }
}

impl UiStepModule for TestUiModule {
    fn get_title(&self) -> String {
        "Test".to_string()
    }

    fn draw_ui(&mut self, ui: &mut Ui) {
        if self.completed {
            ui.label("Done!");
        }else{
            ui.heading("Complete me:");
            ui.add_space(10.0);
            if ui.button("Test").clicked() {
                self.completed = true;
            }
        }

    }

    fn is_complete(&self) -> bool {
        self.completed
    }

    fn reset(&mut self) {
        self.completed = false;
    }
}