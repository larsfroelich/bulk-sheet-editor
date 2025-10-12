use crate::ui_step_modules::{
    SharedState, UiStepModule, column_label_from_index, parse_cell_reference,
};
use calamine::{Data, Reader, open_workbook_auto};
use egui::Ui;
use std::cell::RefCell;
use std::collections::BTreeMap;
use std::fs::File;
use std::io::Write;
use std::path::{Path, PathBuf};
use std::rc::Rc;
use zip::write::FileOptions;
use zip::{CompressionMethod, ZipWriter};

#[derive(Clone)]
struct SheetCell {
    row: u32,
    col: u32,
    value: CellValue,
}

#[derive(Clone)]
struct SheetContent {
    name: String,
    cells: Vec<SheetCell>,
}

#[derive(Clone)]
enum CellValue {
    String(String),
    Number(f64),
    Bool(bool),
}

pub struct BulkCreateModule {
    state: Rc<RefCell<SharedState>>,
    save_path: Option<PathBuf>,
    status_message: Option<String>,
    error_message: Option<String>,
}

impl BulkCreateModule {
    pub fn new(state: Rc<RefCell<SharedState>>) -> Self {
        Self {
            state,
            save_path: None,
            status_message: None,
            error_message: None,
        }
    }

    fn validate_inputs(&self) -> Result<(), String> {
        let state = self.state.borrow();
        if state.csv_rows.is_empty() {
            return Err("Import a CSV file before generating sheets.".to_string());
        }
        if state.odf_path.is_none() || state.selected_sheet.is_none() {
            return Err("Select a template workbook and sheet.".to_string());
        }
        if state
            .cell_mappings
            .iter()
            .all(|mapping| mapping.cell_ref.trim().is_empty())
        {
            return Err("Assign at least one CSV column to a template cell.".to_string());
        }
        Ok(())
    }

    fn generate_and_save(&mut self, path: PathBuf) {
        match self.build_and_write(path.as_path()) {
            Ok(sheet_count) => {
                self.status_message = Some(format!(
                    "Created {} sheet(s) using the template.",
                    sheet_count
                ));
                self.error_message = None;
                self.save_path = Some(path.clone());
                self.state.borrow_mut().last_output_path = Some(path);
            }
            Err(err) => {
                self.error_message = Some(err);
            }
        }
    }

    fn build_and_write(&self, output_path: &Path) -> Result<usize, String> {
        let state = self.state.borrow();
        let template_path = state
            .odf_path
            .clone()
            .ok_or_else(|| "Template workbook missing".to_string())?;
        let template_sheet_name = state
            .selected_sheet
            .clone()
            .ok_or_else(|| "Template sheet missing".to_string())?;
        let mappings = state
            .cell_mappings
            .iter()
            .filter(|mapping| !mapping.cell_ref.trim().is_empty())
            .cloned()
            .collect::<Vec<_>>();
        let rows = state.csv_rows.clone();
        drop(state);

        if rows.is_empty() {
            return Err("CSV file does not contain data rows.".to_string());
        }
        if mappings.is_empty() {
            return Err("No column mappings configured.".to_string());
        }

        let template_cells = load_template_cells(&template_path, &template_sheet_name)?;
        let mut sheets = Vec::new();
        for (row_index, row_values) in rows.iter().enumerate() {
            let mut cell_map = template_cells.clone();
            for mapping in &mappings {
                if let Some((row, col)) = parse_cell_reference(&mapping.cell_ref)
                    && let Some(value) = row_values.get(mapping.column_index)
                {
                    cell_map.insert((row, col), CellValue::String(value.clone()));
                }
            }
            let cells = cell_map
                .into_iter()
                .map(|((row, col), value)| SheetCell { row, col, value })
                .collect();
            sheets.push(SheetContent {
                name: format!("{} {}", template_sheet_name, row_index + 1),
                cells,
            });
        }

        write_xlsx(output_path, &sheets)?;
        Ok(sheets.len())
    }
}

impl UiStepModule for BulkCreateModule {
    fn get_title(&self) -> String {
        "Generate Workbook".to_string()
    }

    fn draw_ui(&mut self, ui: &mut Ui) {
        ui.label("Create sheets for each CSV row and save them as a workbook");

        match self.validate_inputs() {
            Ok(_) => {
                let state = self.state.borrow();
                ui.label(format!("Rows ready for export: {}", state.csv_rows.len()));
                if let Some(path) = &state.odf_path {
                    ui.label(format!("Template file: {}", path.display()));
                }
                if let Some(sheet) = &state.selected_sheet {
                    ui.label(format!("Template sheet: {}", sheet));
                }
            }
            Err(reason) => {
                ui.colored_label(egui::Color32::DARK_RED, reason);
                return;
            }
        }

        ui.add_space(10.0);
        ui.horizontal(|ui| {
            if ui.button("Save as…").clicked()
                && let Some(path) = rfd::FileDialog::new()
                    .add_filter("Excel", &["xlsx"])
                    .set_file_name("bulk_output.xlsx")
                    .save_file()
            {
                self.generate_and_save(path);
            }
            if let Some(path) = &self.save_path {
                ui.label(path.display().to_string());
            }
        });

        if let Some(message) = &self.status_message {
            ui.colored_label(egui::Color32::DARK_GREEN, message);
        }
        if let Some(error) = &self.error_message {
            ui.colored_label(egui::Color32::DARK_RED, error);
        }
    }

    fn is_complete(&self) -> bool {
        self.state.borrow().last_output_path.is_some()
    }

    fn reset(&mut self) {
        self.save_path = None;
        self.status_message = None;
        self.error_message = None;
        self.state.borrow_mut().last_output_path = None;
    }
}

fn load_template_cells(
    path: &Path,
    sheet: &str,
) -> Result<BTreeMap<(u32, u32), CellValue>, String> {
    let mut workbook = open_workbook_auto(path).map_err(|err| err.to_string())?;
    let range = workbook
        .worksheet_range(sheet)
        .map_err(|err| err.to_string())?;

    let mut cells = BTreeMap::new();
    for (row, col, value) in range.cells() {
        if let Some(cell_value) = data_to_value(value) {
            cells.insert((row as u32, col as u32), cell_value);
        }
    }
    Ok(cells)
}

fn data_to_value(data: &Data) -> Option<CellValue> {
    match data {
        Data::String(value) => Some(CellValue::String(value.clone())),
        Data::Float(value) => Some(CellValue::Number(*value)),
        Data::Int(value) => Some(CellValue::Number(*value as f64)),
        Data::Bool(value) => Some(CellValue::Bool(*value)),
        Data::DateTimeIso(value) | Data::DurationIso(value) => {
            Some(CellValue::String(value.clone()))
        }
        Data::DateTime(value) => Some(CellValue::String(value.to_string())),
        Data::Error(err) => Some(CellValue::String(format!("Error: {:?}", err))),
        Data::Empty => None,
    }
}

fn write_xlsx(path: &Path, sheets: &[SheetContent]) -> Result<(), String> {
    let file = File::create(path).map_err(|err| err.to_string())?;
    let mut zip = ZipWriter::new(file);
    let options = FileOptions::default().compression_method(CompressionMethod::Stored);

    write_content_types(&mut zip, options, sheets.len())?;
    write_rels(&mut zip, options)?;
    write_app_doc(&mut zip, options, sheets)?;
    write_core_doc(&mut zip, options)?;
    write_workbook(&mut zip, options, sheets)?;
    write_workbook_rels(&mut zip, options, sheets.len())?;
    write_styles(&mut zip, options)?;
    write_sheets(&mut zip, options, sheets)?;

    zip.finish().map_err(|err| err.to_string()).map(|_| ())
}

fn write_content_types(
    zip: &mut ZipWriter<File>,
    options: FileOptions,
    sheet_count: usize,
) -> Result<(), String> {
    zip.start_file("[Content_Types].xml", options)
        .map_err(|err| err.to_string())?;
    let mut sheet_overrides = String::new();
    for index in 0..sheet_count {
        sheet_overrides.push_str(&format!(
            "    <Override PartName=\"/xl/worksheets/sheet{}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>\n",
            index + 1
        ));
    }
    write!(
        zip,
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\n    <Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\n    <Default Extension=\"xml\" ContentType=\"application/xml\"/>\n    <Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>\n{}    <Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>\n    <Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>\n    <Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>\n</Types>",
        sheet_overrides
    )
    .map_err(|err| err.to_string())
}

fn write_rels(zip: &mut ZipWriter<File>, options: FileOptions) -> Result<(), String> {
    zip.start_file("_rels/.rels", options)
        .map_err(|err| err.to_string())?;
    write!(
        zip,
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n    <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>\n    <Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/>\n    <Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"docProps/app.xml\"/>\n</Relationships>"
    )
    .map_err(|err| err.to_string())
}

fn write_app_doc(
    zip: &mut ZipWriter<File>,
    options: FileOptions,
    sheets: &[SheetContent],
) -> Result<(), String> {
    zip.start_file("docProps/app.xml", options)
        .map_err(|err| err.to_string())?;
    let sheet_count = sheets.len();
    let mut titles = String::new();
    for sheet in sheets {
        titles.push_str(&format!(
            "\n            <vt:lpstr>{}</vt:lpstr>",
            xml_escape(&sheet.name)
        ));
    }
    write!(
        zip,
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">\n    <Application>Bulk Sheet Editor</Application>\n    <DocSecurity>0</DocSecurity>\n    <ScaleCrop>false</ScaleCrop>\n    <HeadingPairs>\n        <vt:vector size=\"2\" baseType=\"variant\">\n            <vt:variant>\n                <vt:lpstr>Worksheets</vt:lpstr>\n            </vt:variant>\n            <vt:variant>\n                <vt:i4>{}</vt:i4>\n            </vt:variant>\n        </vt:vector>\n    </HeadingPairs>\n    <TitlesOfParts>\n        <vt:vector size=\"{}\" baseType=\"lpstr\">{}\n        </vt:vector>\n    </TitlesOfParts>\n    <Company></Company>\n    <LinksUpToDate>false</LinksUpToDate>\n    <SharedDoc>false</SharedDoc>\n    <HyperlinksChanged>false</HyperlinksChanged>\n    <AppVersion>16.0300</AppVersion>\n</Properties>",
        sheet_count,
        sheet_count,
        titles
    )
    .map_err(|err| err.to_string())
}

fn write_core_doc(zip: &mut ZipWriter<File>, options: FileOptions) -> Result<(), String> {
    zip.start_file("docProps/core.xml", options)
        .map_err(|err| err.to_string())?;
    write!(
        zip,
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">\n    <dc:creator>Bulk Sheet Editor</dc:creator>\n    <cp:lastModifiedBy>Bulk Sheet Editor</cp:lastModifiedBy>\n    <dcterms:created xsi:type=\"dcterms:W3CDTF\">2024-01-01T00:00:00Z</dcterms:created>\n    <dcterms:modified xsi:type=\"dcterms:W3CDTF\">2024-01-01T00:00:00Z</dcterms:modified>\n</cp:coreProperties>"
    )
    .map_err(|err| err.to_string())
}

fn write_workbook(
    zip: &mut ZipWriter<File>,
    options: FileOptions,
    sheets: &[SheetContent],
) -> Result<(), String> {
    zip.start_file("xl/workbook.xml", options)
        .map_err(|err| err.to_string())?;
    write!(
        zip,
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n    <sheets>"
    )
    .map_err(|err| err.to_string())?;
    for (index, sheet) in sheets.iter().enumerate() {
        write!(
            zip,
            "\n        <sheet name=\"{}\" sheetId=\"{}\" r:id=\"rId{}\"/>",
            xml_escape(&sheet.name),
            index + 1,
            index + 1
        )
        .map_err(|err| err.to_string())?;
    }
    write!(zip, "\n    </sheets>\n</workbook>").map_err(|err| err.to_string())
}

fn write_workbook_rels(
    zip: &mut ZipWriter<File>,
    options: FileOptions,
    sheet_count: usize,
) -> Result<(), String> {
    zip.start_file("xl/_rels/workbook.xml.rels", options)
        .map_err(|err| err.to_string())?;
    write!(
        zip,
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
    )
    .map_err(|err| err.to_string())?;
    for index in 0..sheet_count {
        write!(
            zip,
            "\n    <Relationship Id=\"rId{}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet{}.xml\"/>",
            index + 1,
            index + 1
        )
        .map_err(|err| err.to_string())?;
    }
    write!(zip, "\n</Relationships>").map_err(|err| err.to_string())
}

fn write_styles(zip: &mut ZipWriter<File>, options: FileOptions) -> Result<(), String> {
    zip.start_file("xl/styles.xml", options)
        .map_err(|err| err.to_string())?;
    write!(
        zip,
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n    <fonts count=\"1\">\n        <font>\n            <sz val=\"11\"/>\n            <color theme=\"1\"/>\n            <name val=\"Calibri\"/>\n            <family val=\"2\"/>\n        </font>\n    </fonts>\n    <fills count=\"2\">\n        <fill><patternFill patternType=\"none\"/></fill>\n        <fill><patternFill patternType=\"gray125\"/></fill>\n    </fills>\n    <borders count=\"1\">\n        <border><left/><right/><top/><bottom/><diagonal/></border>\n    </borders>\n    <cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>\n    <cellXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/></cellXfs>\n    <cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles>\n</styleSheet>"
    )
    .map_err(|err| err.to_string())
}

fn write_sheets(
    zip: &mut ZipWriter<File>,
    options: FileOptions,
    sheets: &[SheetContent],
) -> Result<(), String> {
    for (index, sheet) in sheets.iter().enumerate() {
        zip.start_file(format!("xl/worksheets/sheet{}.xml", index + 1), options)
            .map_err(|err| err.to_string())?;
        write_sheet_content(zip, sheet)?;
    }
    Ok(())
}

fn write_sheet_content(zip: &mut ZipWriter<File>, sheet: &SheetContent) -> Result<(), String> {
    let mut min_row = u32::MAX;
    let mut max_row = 0u32;
    let mut min_col = u32::MAX;
    let mut max_col = 0u32;

    let mut rows: BTreeMap<u32, BTreeMap<u32, &CellValue>> = BTreeMap::new();
    for cell in &sheet.cells {
        min_row = min_row.min(cell.row);
        max_row = max_row.max(cell.row);
        min_col = min_col.min(cell.col);
        max_col = max_col.max(cell.col);
        rows.entry(cell.row)
            .or_default()
            .insert(cell.col, &cell.value);
    }

    let dimension = if sheet.cells.is_empty() {
        "A1".to_string()
    } else {
        format!(
            "{}{}:{}{}",
            column_label_from_index(min_col),
            min_row + 1,
            column_label_from_index(max_col),
            max_row + 1
        )
    };

    write!(
        zip,
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n    <dimension ref=\"{}\"/>\n    <sheetData>",
        dimension
    )
    .map_err(|err| err.to_string())?;

    for (row_index, cells) in rows {
        let row_number = row_index + 1;
        write!(zip, "\n        <row r=\"{}\">", row_number).map_err(|err| err.to_string())?;
        for (col_index, value) in cells {
            let cell_ref = format!("{}{}", column_label_from_index(col_index), row_number);
            match value {
                CellValue::String(value) => {
                    write!(
                        zip,
                        "<c r=\"{}\" t=\"inlineStr\"><is><t>{}</t></is></c>",
                        cell_ref,
                        xml_escape(value)
                    )
                    .map_err(|err| err.to_string())?;
                }
                CellValue::Number(value) => {
                    write!(zip, "<c r=\"{}\"><v>{}</v></c>", cell_ref, value)
                        .map_err(|err| err.to_string())?;
                }
                CellValue::Bool(value) => {
                    let encoded = if *value { 1 } else { 0 };
                    write!(zip, "<c r=\"{}\" t=\"b\"><v>{}</v></c>", cell_ref, encoded)
                        .map_err(|err| err.to_string())?;
                }
            }
        }
        write!(zip, "</row>").map_err(|err| err.to_string())?;
    }

    write!(zip, "\n    </sheetData>\n</worksheet>").map_err(|err| err.to_string())
}

fn xml_escape(value: &str) -> String {
    value
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}
