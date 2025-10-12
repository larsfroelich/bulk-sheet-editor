use crate::ui_step_modules::{
    SharedState, UiStepModule, column_label_from_index, parse_cell_reference,
};
use egui::Ui;
use quick_xml::events::{BytesStart, Event};
use quick_xml::{Reader as XmlReader, Writer as XmlWriter};
use std::cell::RefCell;
use std::collections::BTreeMap;
use std::fs::File;
use std::io::{Read, Write};
use std::path::{Path, PathBuf};
use std::rc::Rc;
use zip::write::FileOptions;
use zip::{CompressionMethod, ZipArchive, ZipWriter};

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

        let context = TemplateContext::load(&template_path, &template_sheet_name)?;
        let mut sheet_exports = Vec::new();
        let mut next_rel_index = context.next_relationship_index;

        for (row_index, row_values) in rows.iter().enumerate() {
            let mut replacements = BTreeMap::new();
            for mapping in &mappings {
                if let Some((row, col)) = parse_cell_reference(&mapping.cell_ref)
                    && let Some(value) = row_values.get(mapping.column_index)
                {
                    let label = format!("{}{}", column_label_from_index(col), row + 1);
                    replacements.insert(label, value.clone());
                }
            }

            let sheet_xml = update_sheet_xml(&context.template_sheet_xml, &replacements)?;
            next_rel_index += 1;
            sheet_exports.push(WorksheetExport {
                name: format!("{} {}", template_sheet_name, row_index + 1),
                relationship_id: format!("rId{}", next_rel_index),
                target: format!("worksheets/sheet{}.xml", row_index + 1),
                sheet_id: (row_index + 1) as u32,
                data: sheet_xml,
                relationship_part: context.template_sheet_relationship.clone(),
            });
        }

        write_workbook_from_template(output_path, &context, &sheet_exports)?;
        Ok(sheet_exports.len())
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

// -- Template workbook helpers -------------------------------------------------

fn write_workbook_from_template(
    path: &Path,
    context: &TemplateContext,
    sheets: &[WorksheetExport],
) -> Result<(), String> {
    let file = File::create(path).map_err(|err| err.to_string())?;
    let mut zip = ZipWriter::new(file);
    let options = FileOptions::default().compression_method(CompressionMethod::Deflated);

    let content_types = build_content_types(&context.content_types_xml, sheets)?;
    let root_rels = build_root_relationships();
    let app_doc = build_app_doc(sheets);
    let core_doc = build_core_doc();
    let workbook_xml = build_workbook_xml(sheets);
    let workbook_rels = build_workbook_rels(&context.preserved_relationships, sheets);

    zip.start_file("[Content_Types].xml", options)
        .map_err(|err| err.to_string())?;
    zip.write_all(&content_types)
        .map_err(|err| err.to_string())?;

    zip.start_file("_rels/.rels", options)
        .map_err(|err| err.to_string())?;
    zip.write_all(&root_rels).map_err(|err| err.to_string())?;

    zip.start_file("docProps/app.xml", options)
        .map_err(|err| err.to_string())?;
    zip.write_all(&app_doc).map_err(|err| err.to_string())?;

    zip.start_file("docProps/core.xml", options)
        .map_err(|err| err.to_string())?;
    zip.write_all(&core_doc).map_err(|err| err.to_string())?;

    zip.start_file("xl/workbook.xml", options)
        .map_err(|err| err.to_string())?;
    zip.write_all(&workbook_xml)
        .map_err(|err| err.to_string())?;

    zip.start_file("xl/_rels/workbook.xml.rels", options)
        .map_err(|err| err.to_string())?;
    zip.write_all(&workbook_rels)
        .map_err(|err| err.to_string())?;

    for sheet in sheets {
        zip.start_file(format!("xl/{}", sheet.target), options)
            .map_err(|err| err.to_string())?;
        zip.write_all(&sheet.data).map_err(|err| err.to_string())?;

        if let Some(rel_data) = &sheet.relationship_part
            && let Some(rel_path) = sheet_relationship_path(&sheet.target)
        {
            zip.start_file(format!("xl/{}", rel_path), options)
                .map_err(|err| err.to_string())?;
            zip.write_all(rel_data).map_err(|err| err.to_string())?;
        }
    }

    for (name, data) in &context.entries {
        if should_skip_entry(name) {
            continue;
        }
        zip.start_file(name, options)
            .map_err(|err| err.to_string())?;
        zip.write_all(data).map_err(|err| err.to_string())?;
    }

    zip.finish().map_err(|err| err.to_string()).map(|_| ())
}

fn update_sheet_xml(
    template: &[u8],
    replacements: &BTreeMap<String, String>,
) -> Result<Vec<u8>, String> {
    if replacements.is_empty() {
        return Ok(template.to_vec());
    }

    let mut reader = XmlReader::from_reader(template);
    reader.trim_text(false);
    let mut writer = XmlWriter::new(Vec::new());
    let mut buffer = Vec::new();
    let mut skip_depth: usize = 0;

    loop {
        match reader
            .read_event_into(&mut buffer)
            .map_err(|err| err.to_string())?
        {
            Event::Eof => break,
            Event::Start(event) => {
                if skip_depth > 0 {
                    skip_depth += 1;
                    continue;
                }

                if event.name().as_ref() == b"c"
                    && let Some(cell_ref) = attribute_value(&event, b"r")
                    && let Some(value) = replacements.get(&cell_ref)
                {
                    let attrs = collect_cell_attributes(&event);
                    write_replaced_cell(&mut writer, &cell_ref, value, &attrs)?;
                    skip_depth = 1;
                    continue;
                }

                writer
                    .write_event(Event::Start(event.into_owned()))
                    .map_err(|err| err.to_string())?;
            }
            Event::Empty(event) => {
                if skip_depth > 0 {
                    continue;
                }

                if event.name().as_ref() == b"c"
                    && let Some(cell_ref) = attribute_value(&event, b"r")
                    && let Some(value) = replacements.get(&cell_ref)
                {
                    let attrs = collect_cell_attributes(&event);
                    write_replaced_cell(&mut writer, &cell_ref, value, &attrs)?;
                    continue;
                }

                writer
                    .write_event(Event::Empty(event.into_owned()))
                    .map_err(|err| err.to_string())?;
            }
            Event::End(event) => {
                if skip_depth > 0 {
                    if skip_depth == 1 && event.name().as_ref() == b"c" {
                        skip_depth = 0;
                    } else if skip_depth > 1 {
                        skip_depth -= 1;
                    }
                    continue;
                }

                writer
                    .write_event(Event::End(event.into_owned()))
                    .map_err(|err| err.to_string())?;
            }
            Event::Text(event) => {
                if skip_depth > 0 {
                    continue;
                }
                writer
                    .write_event(Event::Text(event))
                    .map_err(|err| err.to_string())?;
            }
            Event::Comment(event) => {
                if skip_depth > 0 {
                    continue;
                }
                writer
                    .write_event(Event::Comment(event))
                    .map_err(|err| err.to_string())?;
            }
            Event::CData(event) => {
                if skip_depth > 0 {
                    continue;
                }
                writer
                    .write_event(Event::CData(event))
                    .map_err(|err| err.to_string())?;
            }
            Event::Decl(event) => {
                if skip_depth > 0 {
                    continue;
                }
                writer
                    .write_event(Event::Decl(event.into_owned()))
                    .map_err(|err| err.to_string())?;
            }
            Event::PI(event) => {
                if skip_depth > 0 {
                    continue;
                }
                writer
                    .write_event(Event::PI(event.into_owned()))
                    .map_err(|err| err.to_string())?;
            }
            Event::DocType(event) => {
                if skip_depth > 0 {
                    continue;
                }
                writer
                    .write_event(Event::DocType(event.into_owned()))
                    .map_err(|err| err.to_string())?;
            }
        }
        buffer.clear();
    }

    Ok(writer.into_inner())
}

fn write_replaced_cell(
    writer: &mut XmlWriter<Vec<u8>>,
    reference: &str,
    value: &str,
    attrs: &[(String, String)],
) -> Result<(), String> {
    let mut cell = format!("<c r=\"{}\"", reference);
    for (name, attr_value) in attrs {
        if name == "r" || name == "t" {
            continue;
        }
        cell.push_str(&format!(" {}=\"{}\"", name, attr_value));
    }
    cell.push_str(" t=\"inlineStr\"><is><t>");
    cell.push_str(&xml_escape(value));
    cell.push_str("</t></is></c>");
    writer
        .get_mut()
        .write_all(cell.as_bytes())
        .map_err(|err| err.to_string())
}

fn attribute_value(event: &BytesStart, key: &[u8]) -> Option<String> {
    event
        .attributes()
        .with_checks(false)
        .filter_map(|attr| attr.ok())
        .find(|attr| attr.key.as_ref() == key)
        .map(|attr| String::from_utf8_lossy(attr.value.as_ref()).into_owned())
}

fn collect_cell_attributes(event: &BytesStart) -> Vec<(String, String)> {
    event
        .attributes()
        .with_checks(false)
        .filter_map(|attr| attr.ok())
        .map(|attr| {
            (
                String::from_utf8_lossy(attr.key.as_ref()).into_owned(),
                String::from_utf8_lossy(attr.value.as_ref()).into_owned(),
            )
        })
        .collect()
}

fn build_workbook_xml(sheets: &[WorksheetExport]) -> Vec<u8> {
    let mut xml = String::from(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheets>",
    );
    for sheet in sheets {
        xml.push_str(&format!(
            "<sheet name=\"{}\" sheetId=\"{}\" r:id=\"{}\"/>",
            xml_escape(&sheet.name),
            sheet.sheet_id,
            sheet.relationship_id
        ));
    }
    xml.push_str("</sheets></workbook>");
    xml.into_bytes()
}

fn build_workbook_rels(preserved: &[WorkbookRelationship], sheets: &[WorksheetExport]) -> Vec<u8> {
    let mut xml = String::from(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">",
    );
    for rel in preserved {
        xml.push_str(&format!(
            "<Relationship Id=\"{}\" Type=\"{}\" Target=\"{}\"/>",
            xml_escape(&rel.id),
            xml_escape(&rel.type_attr),
            xml_escape(&rel.target)
        ));
    }
    for sheet in sheets {
        xml.push_str(&format!(
            "<Relationship Id=\"{}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"{}\"/>",
            xml_escape(&sheet.relationship_id),
            xml_escape(&sheet.target)
        ));
    }
    xml.push_str("</Relationships>");
    xml.into_bytes()
}

fn build_content_types(original: &str, sheets: &[WorksheetExport]) -> Result<Vec<u8>, String> {
    let mut reader = XmlReader::from_str(original);
    reader.trim_text(false);
    let mut writer = XmlWriter::new(Vec::new());
    let mut buffer = Vec::new();

    loop {
        match reader
            .read_event_into(&mut buffer)
            .map_err(|err| err.to_string())?
        {
            Event::Eof => break,
            Event::Empty(event) => {
                if event.name().as_ref() == b"Override"
                    && let Some(part_name) = attribute_value(&event, b"PartName")
                    && part_name.contains("/xl/worksheets/")
                {
                    buffer.clear();
                    continue;
                }
                writer
                    .write_event(Event::Empty(event.into_owned()))
                    .map_err(|err| err.to_string())?;
            }
            Event::End(event) => {
                if event.name().as_ref() == b"Types" {
                    for sheet in sheets {
                        let override_line = format!(
                            "\n    <Override PartName=\"/xl/{}\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>",
                            sheet.target
                        );
                        writer
                            .get_mut()
                            .write_all(override_line.as_bytes())
                            .map_err(|err| err.to_string())?;
                    }
                }
                writer
                    .write_event(Event::End(event.into_owned()))
                    .map_err(|err| err.to_string())?;
            }
            other => {
                writer
                    .write_event(other.into_owned())
                    .map_err(|err| err.to_string())?;
            }
        }
        buffer.clear();
    }

    Ok(writer.into_inner())
}

fn build_root_relationships() -> Vec<u8> {
    b"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/><Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/><Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"docProps/app.xml\"/></Relationships>".to_vec()
}

fn build_app_doc(sheets: &[WorksheetExport]) -> Vec<u8> {
    let mut titles = String::new();
    for sheet in sheets {
        titles.push_str(&format!("<vt:lpstr>{}</vt:lpstr>", xml_escape(&sheet.name)));
    }
    format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\"><Application>Bulk Sheet Editor</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop><HeadingPairs><vt:vector size=\"2\" baseType=\"variant\"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>{}</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector size=\"{}\" baseType=\"lpstr\">{}</vt:vector></TitlesOfParts><Company></Company><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>16.0300</AppVersion></Properties>",
        sheets.len(),
        sheets.len(),
        titles
    )
    .into_bytes()
}

fn build_core_doc() -> Vec<u8> {
    b"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"><dc:creator>Bulk Sheet Editor</dc:creator><cp:lastModifiedBy>Bulk Sheet Editor</cp:lastModifiedBy><dcterms:created xsi:type=\"dcterms:W3CDTF\">2024-01-01T00:00:00Z</dcterms:created><dcterms:modified xsi:type=\"dcterms:W3CDTF\">2024-01-01T00:00:00Z</dcterms:modified></cp:coreProperties>".to_vec()
}

fn should_skip_entry(name: &str) -> bool {
    name == "[Content_Types].xml"
        || name == "_rels/.rels"
        || name == "docProps/app.xml"
        || name == "docProps/core.xml"
        || name == "xl/workbook.xml"
        || name == "xl/_rels/workbook.xml.rels"
        || name.starts_with("xl/worksheets/")
}

fn sheet_relationship_path(target: &str) -> Option<String> {
    let (folder, file) = target.rsplit_once('/')?;
    Some(format!("{}/_rels/{}.rels", folder, file))
}

#[derive(Clone)]
struct WorksheetExport {
    name: String,
    relationship_id: String,
    target: String,
    sheet_id: u32,
    data: Vec<u8>,
    relationship_part: Option<Vec<u8>>,
}

#[derive(Clone)]
struct WorkbookRelationship {
    id: String,
    target: String,
    type_attr: String,
}

struct TemplateContext {
    entries: BTreeMap<String, Vec<u8>>,
    content_types_xml: String,
    preserved_relationships: Vec<WorkbookRelationship>,
    next_relationship_index: u32,
    template_sheet_xml: Vec<u8>,
    template_sheet_relationship: Option<Vec<u8>>,
}

impl TemplateContext {
    fn load(path: &Path, sheet_name: &str) -> Result<Self, String> {
        let file = File::open(path).map_err(|err| err.to_string())?;
        let mut archive = ZipArchive::new(file).map_err(|err| err.to_string())?;
        let mut entries = BTreeMap::new();
        for index in 0..archive.len() {
            let mut file = archive.by_index(index).map_err(|err| err.to_string())?;
            if !file.is_file() {
                continue;
            }
            let mut data = Vec::new();
            file.read_to_end(&mut data).map_err(|err| err.to_string())?;
            entries.insert(file.name().to_string(), data);
        }

        let content_types_xml = entries
            .get("[Content_Types].xml")
            .ok_or_else(|| "Workbook content types missing".to_string())?
            .clone();
        let content_types_xml =
            String::from_utf8(content_types_xml).map_err(|err| err.to_string())?;

        let workbook_xml = entries
            .get("xl/workbook.xml")
            .ok_or_else(|| "Workbook definition missing".to_string())?
            .clone();
        let workbook_xml = String::from_utf8(workbook_xml).map_err(|err| err.to_string())?;

        let workbook_rels = entries
            .get("xl/_rels/workbook.xml.rels")
            .ok_or_else(|| "Workbook relationships missing".to_string())?
            .clone();
        let workbook_rels = String::from_utf8(workbook_rels).map_err(|err| err.to_string())?;

        let template_rel_id = parse_sheet_mapping(&workbook_xml, sheet_name)?;
        let (template_target, preserved_relationships, next_relationship_index) =
            parse_workbook_relationships(&workbook_rels, &template_rel_id)?;

        let sheet_entry = format!("xl/{}", template_target);
        let template_sheet_xml = entries
            .get(&sheet_entry)
            .ok_or_else(|| "Template sheet XML missing".to_string())?
            .clone();

        let relationship_part = sheet_relationship_path(&template_target)
            .and_then(|path| entries.get(&format!("xl/{}", path)).cloned());

        Ok(Self {
            entries,
            content_types_xml,
            preserved_relationships,
            next_relationship_index,
            template_sheet_xml,
            template_sheet_relationship: relationship_part,
        })
    }
}

fn parse_sheet_mapping(workbook_xml: &str, sheet_name: &str) -> Result<String, String> {
    let mut reader = XmlReader::from_str(workbook_xml);
    reader.trim_text(true);
    let mut buffer = Vec::new();
    let mut template_rel = None;

    loop {
        match reader
            .read_event_into(&mut buffer)
            .map_err(|err| err.to_string())?
        {
            Event::Eof => break,
            Event::Empty(event) => {
                if event.name().as_ref() == b"sheet" {
                    let mut name = None;
                    let mut rel_id = None;
                    for attr in event.attributes().with_checks(false) {
                        let attr = attr.map_err(|err| err.to_string())?;
                        let key = attr.key.as_ref();
                        let value = String::from_utf8_lossy(attr.value.as_ref()).into_owned();
                        if key == b"name" {
                            name = Some(value);
                        } else if key == b"r:id" {
                            rel_id = Some(value);
                        }
                    }

                    if name.as_deref() == Some(sheet_name)
                        && let Some(rel) = rel_id
                    {
                        template_rel = Some(rel);
                    }
                }
            }
            _ => {}
        }
        buffer.clear();
    }

    template_rel.ok_or_else(|| "Template sheet not found".to_string())
}

fn parse_workbook_relationships(
    xml: &str,
    template_rel_id: &str,
) -> Result<(String, Vec<WorkbookRelationship>, u32), String> {
    let mut reader = XmlReader::from_str(xml);
    reader.trim_text(true);
    let mut buffer = Vec::new();
    let mut template_target = None;
    let mut preserved = Vec::new();
    let mut max_id = 0u32;

    loop {
        match reader
            .read_event_into(&mut buffer)
            .map_err(|err| err.to_string())?
        {
            Event::Eof => break,
            Event::Empty(event) => {
                if event.name().as_ref() == b"Relationship" {
                    let mut id = None;
                    let mut target = None;
                    let mut kind = None;
                    for attr in event.attributes().with_checks(false) {
                        let attr = attr.map_err(|err| err.to_string())?;
                        let key = attr.key.as_ref();
                        let value = String::from_utf8_lossy(attr.value.as_ref()).into_owned();
                        if key == b"Id" {
                            id = Some(value.clone());
                            if let Some(suffix) = value.strip_prefix("rId")
                                && let Ok(number) = suffix.parse::<u32>()
                            {
                                max_id = max_id.max(number);
                            }
                        } else if key == b"Target" {
                            target = Some(value);
                        } else if key == b"Type" {
                            kind = Some(value);
                        }
                    }

                    let id = id.ok_or_else(|| "Relationship id missing".to_string())?;
                    let target = target.ok_or_else(|| "Relationship target missing".to_string())?;
                    let kind = kind.ok_or_else(|| "Relationship type missing".to_string())?;

                    if kind.ends_with("/worksheet") {
                        if id == template_rel_id {
                            template_target = Some(target);
                        }
                    } else {
                        preserved.push(WorkbookRelationship {
                            id,
                            target,
                            type_attr: kind,
                        });
                    }
                }
            }
            _ => {}
        }
        buffer.clear();
    }

    let target =
        template_target.ok_or_else(|| "Template sheet relationship missing".to_string())?;
    Ok((target, preserved, max_id))
}

fn xml_escape(value: &str) -> String {
    value
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}
