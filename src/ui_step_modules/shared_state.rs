use std::collections::HashMap;
use std::path::PathBuf;

#[derive(Default)]
pub struct SharedState {
    pub csv_path: Option<PathBuf>,
    pub csv_has_headers: bool,
    pub csv_headers: Vec<String>,
    pub csv_rows: Vec<Vec<String>>,
    pub csv_preview: Vec<ColumnPreview>,
    pub odf_path: Option<PathBuf>,
    pub odf_sheet_names: Vec<String>,
    pub selected_sheet: Option<String>,
    pub template_cell_values: HashMap<String, String>,
    pub cell_mappings: Vec<CellMapping>,
    pub last_output_path: Option<PathBuf>,
}

impl SharedState {
    pub fn reset_csv(&mut self) {
        self.csv_path = None;
        self.csv_headers.clear();
        self.csv_rows.clear();
        self.csv_preview.clear();
        self.csv_has_headers = false;
        self.cell_mappings.clear();
    }

    pub fn reset_template(&mut self) {
        self.odf_path = None;
        self.odf_sheet_names.clear();
        self.selected_sheet = None;
        self.template_cell_values.clear();
        for mapping in &mut self.cell_mappings {
            mapping.cell_ref.clear();
        }
    }

    pub fn ensure_cell_mappings(&mut self) {
        if self.cell_mappings.len() > self.csv_headers.len() {
            self.cell_mappings.truncate(self.csv_headers.len());
        }
        while self.cell_mappings.len() < self.csv_headers.len() {
            let column_index = self.cell_mappings.len();
            self.cell_mappings
                .push(CellMapping::new(column_index, String::new()));
        }
    }
}

#[derive(Clone, Default)]
pub struct ColumnPreview {
    pub index: usize,
    pub header: String,
    pub samples: Vec<String>,
}

#[derive(Clone, Default)]
pub struct CellMapping {
    pub column_index: usize,
    pub cell_ref: String,
}

impl CellMapping {
    pub fn new<S: Into<String>>(column_index: usize, cell_ref: S) -> Self {
        Self {
            column_index,
            cell_ref: cell_ref.into(),
        }
    }
}

pub fn parse_cell_reference(cell: &str) -> Option<(u32, u32)> {
    if cell.is_empty() {
        return None;
    }

    let mut col_index: u32 = 0;
    let mut row_part = String::new();
    for ch in cell.chars() {
        if ch.is_ascii_alphabetic() {
            col_index = col_index * 26 + u32::from((ch.to_ascii_uppercase() as u8) - b'A' + 1);
        } else if ch.is_ascii_digit() {
            row_part.push(ch);
        } else {
            return None;
        }
    }
    if col_index == 0 || row_part.is_empty() {
        return None;
    }
    let row = row_part.parse::<u32>().ok()?.saturating_sub(1);
    Some((row, col_index - 1))
}

pub fn column_label_from_index(index: u32) -> String {
    let mut idx = index + 1;
    let mut label = String::new();
    while idx > 0 {
        let rem = ((idx - 1) % 26) as u8;
        label.insert(0, char::from(b'A' + rem));
        idx = (idx - 1) / 26;
    }
    label
}
