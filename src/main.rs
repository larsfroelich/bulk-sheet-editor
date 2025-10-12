mod csv_loader;
mod ui_step_modules;

extern crate alloc;

use crate::ui_step_modules::{CsvImportModule, TestUiModule, UiStepModule};
use alloc::string::String;
use catppuccin_egui::{LATTE, MOCHA, set_theme};
use egui::{Align, Color32, FontId, Layout, RichText, Vec2};

#[derive(Default)]
pub struct BulkSheetEditorApp {
    dark_theme: bool,
    ui_step_modules: Vec<Box<dyn UiStepModule>>,
}

impl BulkSheetEditorApp {
    fn new() -> Self {
        Self {
            dark_theme: false,
            ui_step_modules: vec![
                Box::new(TestUiModule::new()),
                Box::new(TestUiModule::new()),
                Box::new(TestUiModule::new()),
                Box::new(CsvImportModule::new()),
            ],
        }
    }
}

const VERSION: &str = env!("CARGO_PKG_VERSION");
impl eframe::App for BulkSheetEditorApp {
    fn update(&mut self, ctx: &egui::Context, _frame: &mut eframe::Frame) {
        egui::CentralPanel::default().show(ctx, |ui| {
            ui.allocate_ui_with_layout(
                Vec2::new(0.0, 0.0),
                Layout::left_to_right(Align::BOTTOM),
                |ui| {
                    ui.label(
                        RichText::new("Bulk Sheet Editor")
                            .font(FontId::proportional(40.0))
                            .line_height(Some(40.0)),
                    );
                    ui.label(
                        RichText::from(String::from(" v") + VERSION)
                            .font(FontId::proportional(20.0)),
                    );
                    if ui
                        .button(if !self.dark_theme { "🔆" } else { "🌙 " })
                        .clicked()
                    {
                        self.dark_theme = !self.dark_theme;
                        set_theme(ctx, if self.dark_theme { MOCHA } else { LATTE });
                    }
                },
            );
            ui.separator();
            ui.add_space(25.0);

            // scrollable steps
            egui::ScrollArea::vertical()
                .auto_shrink(false)
                .show(ui, |ui| {
                    for step_nr in 0..self.ui_step_modules.len() {
                        let is_previous_step_complete = step_nr == 0
                            || self.ui_step_modules[step_nr.saturating_sub(1)].is_complete();
                        let is_current_step = !self.ui_step_modules[step_nr].is_complete()
                            && is_previous_step_complete;
                        ui.horizontal_wrapped(|ui| {
                            ui.set_min_height(32.0);
                            if self.ui_step_modules[step_nr].is_complete() {
                                if ui
                                    .button(RichText::new("✅").color(Color32::DARK_GREEN))
                                    .on_hover_text("reset to this step")
                                    .clicked()
                                {
                                    for remaining_step_nr in step_nr..self.ui_step_modules.len() {
                                        self.ui_step_modules[remaining_step_nr].reset();
                                    }
                                }
                            } else if is_previous_step_complete {
                                ui.label("➡");
                            } else {
                                ui.label("...").on_hover_text("complete other tasks first");
                            }
                            let step_text = RichText::new(format!(
                                "Step {}: {}",
                                step_nr + 1,
                                self.ui_step_modules[step_nr].get_title()
                            ))
                            .font(FontId::proportional(25.0));
                            ui.label(if is_current_step {
                                step_text.underline()
                            } else {
                                step_text
                            });
                        });

                        if is_previous_step_complete {
                            ui.indent(20u32, |ui| {
                                self.ui_step_modules[step_nr].draw_ui(ui);
                            });
                        }
                        ui.add_space(15.0);
                    }
                })
        });
    }
}

fn main() -> eframe::Result {
    // init env logger
    env_logger::init();

    // run the app
    eframe::run_native(
        "File Kraken",
        eframe::NativeOptions {
            viewport: egui::ViewportBuilder::default()
                .with_maximized(false)
                .with_min_inner_size(Vec2::from([1200.0, 800.0])),
            ..Default::default()
        },
        Box::new(|cc| {
            cc.egui_ctx.set_zoom_factor(1.6);
            set_theme(&cc.egui_ctx, LATTE);
            Ok(Box::from(BulkSheetEditorApp::new()))
        }),
    )
}
