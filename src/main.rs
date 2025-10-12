extern crate alloc;

use alloc::string::String;
use catppuccin_egui::{set_theme, LATTE, MOCHA};
use egui::{Align, FontId, Layout, RichText, Vec2};

#[derive(Default)]
pub struct BulkSheetEditorApp {
    dark_theme: bool,
}

impl BulkSheetEditorApp {
    fn new()->Self{
        Self { dark_theme: false }
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
            viewport: egui::ViewportBuilder::default().with_maximized(false),
            ..Default::default()
        },
        Box::new(|cc| {
            cc.egui_ctx.set_zoom_factor(1.6);
            set_theme(&cc.egui_ctx, LATTE);
            Ok(Box::from(BulkSheetEditorApp::new()))
        }),
    )
}
