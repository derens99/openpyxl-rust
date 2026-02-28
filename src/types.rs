use std::collections::HashMap;

#[derive(Clone, Debug)]
pub(crate) enum CellData {
    String(String),
    Number(f64),
    Boolean(bool),
    Formula(String),
    DateTime(f64, u8), // (serial, kind: 0=date, 1=time, 2=datetime)
    RichText(String),   // JSON-serialized rich text segments
    Empty,
}

#[derive(Clone, Debug, Default)]
pub(crate) struct CellFormat {
    pub(crate) font_bold: bool,
    pub(crate) font_italic: bool,
    pub(crate) font_name: Option<String>,
    pub(crate) font_size: Option<f64>,
    pub(crate) font_color: Option<String>,
    pub(crate) font_underline: Option<u8>, // 0=none, 1=single, 2=double
    pub(crate) font_strikethrough: bool,
    pub(crate) font_vert_align: Option<u8>, // 1=superscript, 2=subscript
    pub(crate) number_format: Option<String>,
    pub(crate) align_horizontal: Option<u8>, // 1=left,2=center,3=right,4=fill,5=justify,6=centerAcross,7=distributed
    pub(crate) align_vertical: Option<u8>,   // 1=top,2=center,3=bottom,4=justify,5=distributed
    pub(crate) align_wrap_text: bool,
    pub(crate) align_shrink_to_fit: bool,
    pub(crate) align_indent: u8,
    pub(crate) align_text_rotation: i16,
    pub(crate) fill_type: Option<u8>, // 1=solid,2=darkGray,3=mediumGray,4=lightGray,5=gray125,6=gray0625
    pub(crate) fill_start_color: Option<String>,
    pub(crate) fill_end_color: Option<String>,
    pub(crate) border_left_style: Option<u8>,
    pub(crate) border_left_color: Option<String>,
    pub(crate) border_right_style: Option<u8>,
    pub(crate) border_right_color: Option<String>,
    pub(crate) border_top_style: Option<u8>,
    pub(crate) border_top_color: Option<String>,
    pub(crate) border_bottom_style: Option<u8>,
    pub(crate) border_bottom_color: Option<String>,
    pub(crate) border_diagonal_style: Option<u8>,
    pub(crate) border_diagonal_color: Option<String>,
    pub(crate) border_diagonal_up: bool,
    pub(crate) border_diagonal_down: bool,
    pub(crate) protection_locked: Option<bool>,
    pub(crate) protection_hidden: Option<bool>,
}

#[derive(Clone, Debug)]
pub(crate) struct CellValue {
    pub(crate) value: CellData,
    pub(crate) format: CellFormat,
}

#[derive(Clone, Debug)]
pub(crate) struct SheetData {
    pub(crate) title: String,
    pub(crate) cells: HashMap<(u32, u16), CellValue>, // (row, col) -> CellValue (0-based)
    pub(crate) column_widths: HashMap<u16, f64>,      // col (0-based) -> width
    pub(crate) row_heights: HashMap<u32, f64>,        // row (0-based) -> height
    pub(crate) freeze_panes: Option<(u32, u16)>,      // (row, col) 0-based
    pub(crate) merged_ranges: Vec<(u32, u16, u32, u16)>, // (r1, c1, r2, c2) 0-based
    #[allow(clippy::type_complexity)]
    pub(crate) hyperlinks: Vec<(u32, u16, String, Option<String>, Option<String>)>, // (row, col, url, text, tooltip)
    pub(crate) notes: Vec<(u32, u16, String, Option<String>)>, // (row, col, text, author)
    pub(crate) autofilter: Option<(u32, u16, u32, u16)>,       // (r1, c1, r2, c2) 0-based
    pub(crate) autofilter_columns: Vec<String>,  // JSON-serialized per-column filters
    pub(crate) protection_json: Option<String>,
    pub(crate) page_setup_json: Option<String>,
    #[allow(clippy::type_complexity)]
    pub(crate) images: Vec<(u32, u16, Vec<u8>, Option<f64>, Option<f64>)>, // (row, col, image_data, scale_width, scale_height)
    pub(crate) data_validations: Vec<String>,
    pub(crate) conditional_formats: Vec<String>,
    pub(crate) tables: Vec<String>,
    pub(crate) charts: Vec<String>,
    pub(crate) visibility: u8,                // 0=visible, 1=hidden, 2=veryHidden
    pub(crate) hidden_rows: Vec<u32>,         // 0-based row indices
    pub(crate) hidden_cols: Vec<u16>,         // 0-based col indices
    pub(crate) zoom: Option<u16>,
    pub(crate) show_gridlines: Option<bool>,
    pub(crate) autofit: bool,
    pub(crate) row_breaks: Vec<u32>,
    pub(crate) col_breaks: Vec<u16>,
    pub(crate) row_outline_levels: Vec<(u32, u8)>,  // (row, level)
    pub(crate) col_outline_levels: Vec<(u16, u8)>,  // (col, level)
    pub(crate) min_row: Option<u32>,
    pub(crate) max_row: Option<u32>,
    pub(crate) min_col: Option<u16>,
    pub(crate) max_col: Option<u16>,
    pub(crate) append_row: u32,
}

impl SheetData {
    pub(crate) fn new(title: String) -> Self {
        SheetData {
            title,
            cells: HashMap::new(),
            column_widths: HashMap::new(),
            row_heights: HashMap::new(),
            freeze_panes: None,
            merged_ranges: Vec::new(),
            hyperlinks: Vec::new(),
            notes: Vec::new(),
            autofilter: None,
            autofilter_columns: Vec::new(),
            protection_json: None,
            page_setup_json: None,
            images: Vec::new(),
            data_validations: Vec::new(),
            conditional_formats: Vec::new(),
            tables: Vec::new(),
            charts: Vec::new(),
            visibility: 0,
            hidden_rows: Vec::new(),
            hidden_cols: Vec::new(),
            zoom: None,
            show_gridlines: None,
            autofit: false,
            row_breaks: Vec::new(),
            col_breaks: Vec::new(),
            row_outline_levels: Vec::new(),
            col_outline_levels: Vec::new(),
            min_row: None,
            max_row: None,
            min_col: None,
            max_col: None,
            append_row: 0,
        }
    }

    pub(crate) fn track_cell(&mut self, row: u32, col: u16) {
        match self.min_row {
            None => {
                self.min_row = Some(row);
                self.max_row = Some(row);
                self.min_col = Some(col);
                self.max_col = Some(col);
            }
            Some(mn) => {
                if row < mn {
                    self.min_row = Some(row);
                }
                if row > self.max_row.unwrap() {
                    self.max_row = Some(row);
                }
                if col < self.min_col.unwrap() {
                    self.min_col = Some(col);
                }
                if col > self.max_col.unwrap() {
                    self.max_col = Some(col);
                }
            }
        }
    }

    pub(crate) fn recompute_dimensions(&mut self) {
        if self.cells.is_empty() {
            self.min_row = None;
            self.max_row = None;
            self.min_col = None;
            self.max_col = None;
        } else {
            let mut mn_r = u32::MAX;
            let mut mx_r = 0u32;
            let mut mn_c = u16::MAX;
            let mut mx_c = 0u16;
            for &(r, c) in self.cells.keys() {
                if r < mn_r {
                    mn_r = r;
                }
                if r > mx_r {
                    mx_r = r;
                }
                if c < mn_c {
                    mn_c = c;
                }
                if c > mx_c {
                    mx_c = c;
                }
            }
            self.min_row = Some(mn_r);
            self.max_row = Some(mx_r);
            self.min_col = Some(mn_c);
            self.max_col = Some(mx_c);
        }
    }
}
