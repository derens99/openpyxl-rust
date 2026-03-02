use crate::types::CellFormat;
use rust_xlsxwriter::Format;

pub(crate) fn parse_border_style_str(s: &str) -> rust_xlsxwriter::FormatBorder {
    match s {
        "thin" => rust_xlsxwriter::FormatBorder::Thin,
        "medium" => rust_xlsxwriter::FormatBorder::Medium,
        "thick" => rust_xlsxwriter::FormatBorder::Thick,
        "dashed" => rust_xlsxwriter::FormatBorder::Dashed,
        "dotted" => rust_xlsxwriter::FormatBorder::Dotted,
        "double" => rust_xlsxwriter::FormatBorder::Double,
        "hair" => rust_xlsxwriter::FormatBorder::Hair,
        "mediumDashed" => rust_xlsxwriter::FormatBorder::MediumDashed,
        "dashDot" => rust_xlsxwriter::FormatBorder::DashDot,
        "mediumDashDot" => rust_xlsxwriter::FormatBorder::MediumDashDot,
        "dashDotDot" => rust_xlsxwriter::FormatBorder::DashDotDot,
        "mediumDashDotDot" => rust_xlsxwriter::FormatBorder::MediumDashDotDot,
        "slantDashDot" => rust_xlsxwriter::FormatBorder::SlantDashDot,
        _ => rust_xlsxwriter::FormatBorder::Thin,
    }
}

pub(crate) fn parse_color_str(c: &str) -> Option<rust_xlsxwriter::Color> {
    u32::from_str_radix(c, 16)
        .ok()
        .map(rust_xlsxwriter::Color::from)
}

pub(crate) fn build_format_from_json(json_str: &str) -> Result<Format, String> {
    let val: serde_json::Value =
        serde_json::from_str(json_str).map_err(|e| format!("JSON parse error: {}", e))?;
    let obj = val.as_object().ok_or("Expected JSON object")?;
    let mut fmt = Format::new();

    // Font
    if let Some(font) = obj.get("font").and_then(|v| v.as_object()) {
        if let Some(bold) = font.get("bold").and_then(|v| v.as_bool()) {
            if bold {
                fmt = fmt.set_bold();
            }
        }
        if let Some(italic) = font.get("italic").and_then(|v| v.as_bool()) {
            if italic {
                fmt = fmt.set_italic();
            }
        }
        if let Some(ul) = font.get("underline").and_then(|v| v.as_str()) {
            if ul == "single" {
                fmt = fmt.set_underline(rust_xlsxwriter::FormatUnderline::Single);
            } else if ul == "double" {
                fmt = fmt.set_underline(rust_xlsxwriter::FormatUnderline::Double);
            }
        }
        if let Some(name) = font.get("name").and_then(|v| v.as_str()) {
            fmt = fmt.set_font_name(name);
        }
        if let Some(size) = font.get("size").and_then(|v| v.as_f64()) {
            fmt = fmt.set_font_size(size);
        }
        if let Some(color) = font.get("color").and_then(|v| v.as_str()) {
            if let Some(clr) = parse_color_str(color) {
                fmt = fmt.set_font_color(clr);
            }
        }
        if let Some(st) = font.get("strikethrough").and_then(|v| v.as_bool()) {
            if st {
                fmt = fmt.set_font_strikethrough();
            }
        }
        if let Some(va) = font.get("vertAlign").and_then(|v| v.as_str()) {
            match va {
                "superscript" => {
                    fmt = fmt.set_font_script(rust_xlsxwriter::FormatScript::Superscript);
                }
                "subscript" => {
                    fmt = fmt.set_font_script(rust_xlsxwriter::FormatScript::Subscript);
                }
                _ => {}
            }
        }
    }

    // Alignment
    if let Some(align) = obj.get("alignment").and_then(|v| v.as_object()) {
        if let Some(h) = align.get("horizontal").and_then(|v| v.as_str()) {
            let a = match h {
                "center" => rust_xlsxwriter::FormatAlign::Center,
                "right" => rust_xlsxwriter::FormatAlign::Right,
                "left" => rust_xlsxwriter::FormatAlign::Left,
                "fill" => rust_xlsxwriter::FormatAlign::Fill,
                "justify" => rust_xlsxwriter::FormatAlign::Justify,
                "centerContinuous" | "center_continuous" => {
                    rust_xlsxwriter::FormatAlign::CenterAcross
                }
                "distributed" => rust_xlsxwriter::FormatAlign::Distributed,
                _ => rust_xlsxwriter::FormatAlign::General,
            };
            fmt = fmt.set_align(a);
        }
        if let Some(v) = align.get("vertical").and_then(|v| v.as_str()) {
            let a = match v {
                "center" => rust_xlsxwriter::FormatAlign::VerticalCenter,
                "top" => rust_xlsxwriter::FormatAlign::Top,
                "bottom" => rust_xlsxwriter::FormatAlign::Bottom,
                "justify" => rust_xlsxwriter::FormatAlign::VerticalJustify,
                "distributed" => rust_xlsxwriter::FormatAlign::VerticalDistributed,
                _ => rust_xlsxwriter::FormatAlign::Bottom,
            };
            fmt = fmt.set_align(a);
        }
        if let Some(wt) = align.get("wrap_text").and_then(|v| v.as_bool()) {
            if wt {
                fmt = fmt.set_text_wrap();
            }
        }
        if let Some(sf) = align.get("shrink_to_fit").and_then(|v| v.as_bool()) {
            if sf {
                fmt = fmt.set_shrink();
            }
        }
        if let Some(indent) = align.get("indent").and_then(|v| v.as_u64()) {
            if indent > 0 {
                fmt = fmt.set_indent(indent as u8);
            }
        }
        if let Some(rot) = align.get("text_rotation").and_then(|v| v.as_i64()) {
            if rot != 0 {
                fmt = fmt.set_rotation(rot as i16);
            }
        }
    }

    // Border
    if let Some(border) = obj.get("border").and_then(|v| v.as_object()) {
        if let Some(left) = border.get("left").and_then(|v| v.as_object()) {
            if let Some(style) = left.get("style").and_then(|v| v.as_str()) {
                fmt = fmt.set_border_left(parse_border_style_str(style));
            }
            if let Some(color) = left.get("color").and_then(|v| v.as_str()) {
                if let Some(clr) = parse_color_str(color) {
                    fmt = fmt.set_border_left_color(clr);
                }
            }
        }
        if let Some(right) = border.get("right").and_then(|v| v.as_object()) {
            if let Some(style) = right.get("style").and_then(|v| v.as_str()) {
                fmt = fmt.set_border_right(parse_border_style_str(style));
            }
            if let Some(color) = right.get("color").and_then(|v| v.as_str()) {
                if let Some(clr) = parse_color_str(color) {
                    fmt = fmt.set_border_right_color(clr);
                }
            }
        }
        if let Some(top) = border.get("top").and_then(|v| v.as_object()) {
            if let Some(style) = top.get("style").and_then(|v| v.as_str()) {
                fmt = fmt.set_border_top(parse_border_style_str(style));
            }
            if let Some(color) = top.get("color").and_then(|v| v.as_str()) {
                if let Some(clr) = parse_color_str(color) {
                    fmt = fmt.set_border_top_color(clr);
                }
            }
        }
        if let Some(bottom) = border.get("bottom").and_then(|v| v.as_object()) {
            if let Some(style) = bottom.get("style").and_then(|v| v.as_str()) {
                fmt = fmt.set_border_bottom(parse_border_style_str(style));
            }
            if let Some(color) = bottom.get("color").and_then(|v| v.as_str()) {
                if let Some(clr) = parse_color_str(color) {
                    fmt = fmt.set_border_bottom_color(clr);
                }
            }
        }
        if let Some(diag) = border.get("diagonal").and_then(|v| v.as_object()) {
            if let Some(style) = diag.get("style").and_then(|v| v.as_str()) {
                fmt = fmt.set_border_diagonal(parse_border_style_str(style));
            }
            if let Some(color) = diag.get("color").and_then(|v| v.as_str()) {
                if let Some(clr) = parse_color_str(color) {
                    fmt = fmt.set_border_diagonal_color(clr);
                }
            }
            let diag_up = diag
                .get("diagonalUp")
                .and_then(|v| v.as_bool())
                .unwrap_or(false);
            let diag_down = diag
                .get("diagonalDown")
                .and_then(|v| v.as_bool())
                .unwrap_or(false);
            let diag_type = match (diag_up, diag_down) {
                (true, true) => rust_xlsxwriter::FormatDiagonalBorder::BorderUpDown,
                (true, false) => rust_xlsxwriter::FormatDiagonalBorder::BorderUp,
                (false, true) => rust_xlsxwriter::FormatDiagonalBorder::BorderDown,
                (false, false) => rust_xlsxwriter::FormatDiagonalBorder::BorderUp,
            };
            fmt = fmt.set_border_diagonal_type(diag_type);
        }
    }

    // Fill
    if let Some(fill) = obj.get("fill").and_then(|v| v.as_object()) {
        if let Some(ft) = fill.get("fill_type").and_then(|v| v.as_str()) {
            let pattern = match ft {
                "solid" => rust_xlsxwriter::FormatPattern::Solid,
                "darkGray" => rust_xlsxwriter::FormatPattern::DarkGray,
                "mediumGray" => rust_xlsxwriter::FormatPattern::MediumGray,
                "lightGray" => rust_xlsxwriter::FormatPattern::LightGray,
                "gray125" => rust_xlsxwriter::FormatPattern::Gray125,
                "gray0625" => rust_xlsxwriter::FormatPattern::Gray0625,
                _ => rust_xlsxwriter::FormatPattern::Solid,
            };
            fmt = fmt.set_pattern(pattern);
        }
        if let Some(sc) = fill.get("start_color").and_then(|v| v.as_str()) {
            if let Some(clr) = parse_color_str(sc) {
                fmt = fmt.set_background_color(clr);
            }
        }
        if let Some(ec) = fill.get("end_color").and_then(|v| v.as_str()) {
            if let Some(clr) = parse_color_str(ec) {
                fmt = fmt.set_foreground_color(clr);
            }
        }
    }

    // Number format
    if let Some(nf) = obj.get("number_format").and_then(|v| v.as_str()) {
        if nf != "General" {
            fmt = fmt.set_num_format(nf);
        }
    }

    Ok(fmt)
}

// =====================================================================
// CellFormat -> rust_xlsxwriter::Format conversion
// =====================================================================

pub(crate) fn border_style_from_u8(v: u8) -> rust_xlsxwriter::FormatBorder {
    match v {
        1 => rust_xlsxwriter::FormatBorder::Thin,
        2 => rust_xlsxwriter::FormatBorder::Medium,
        3 => rust_xlsxwriter::FormatBorder::Thick,
        4 => rust_xlsxwriter::FormatBorder::Dashed,
        5 => rust_xlsxwriter::FormatBorder::Dotted,
        6 => rust_xlsxwriter::FormatBorder::Double,
        7 => rust_xlsxwriter::FormatBorder::Hair,
        8 => rust_xlsxwriter::FormatBorder::MediumDashed,
        9 => rust_xlsxwriter::FormatBorder::DashDot,
        10 => rust_xlsxwriter::FormatBorder::MediumDashDot,
        11 => rust_xlsxwriter::FormatBorder::DashDotDot,
        12 => rust_xlsxwriter::FormatBorder::MediumDashDotDot,
        13 => rust_xlsxwriter::FormatBorder::SlantDashDot,
        _ => rust_xlsxwriter::FormatBorder::Thin,
    }
}

pub(crate) fn fill_pattern_from_u8(v: u8) -> rust_xlsxwriter::FormatPattern {
    match v {
        1 => rust_xlsxwriter::FormatPattern::Solid,
        2 => rust_xlsxwriter::FormatPattern::DarkGray,
        3 => rust_xlsxwriter::FormatPattern::MediumGray,
        4 => rust_xlsxwriter::FormatPattern::LightGray,
        5 => rust_xlsxwriter::FormatPattern::Gray125,
        6 => rust_xlsxwriter::FormatPattern::Gray0625,
        _ => rust_xlsxwriter::FormatPattern::Solid,
    }
}

pub(crate) fn cell_format_to_xlsx_format(cf: &CellFormat) -> (Format, bool) {
    let mut fmt = Format::new();
    let mut has_format = false;

    // Font
    if cf.font_bold {
        fmt = fmt.set_bold();
        has_format = true;
    }
    if cf.font_italic {
        fmt = fmt.set_italic();
        has_format = true;
    }
    if let Some(ref name) = cf.font_name {
        fmt = fmt.set_font_name(name);
        has_format = true;
    }
    if let Some(size) = cf.font_size {
        fmt = fmt.set_font_size(size);
        has_format = true;
    }
    if let Some(ref color) = cf.font_color {
        if let Some(clr) = parse_color_str(color) {
            fmt = fmt.set_font_color(clr);
            has_format = true;
        }
    }
    if let Some(ul) = cf.font_underline {
        match ul {
            1 => {
                fmt = fmt.set_underline(rust_xlsxwriter::FormatUnderline::Single);
                has_format = true;
            }
            2 => {
                fmt = fmt.set_underline(rust_xlsxwriter::FormatUnderline::Double);
                has_format = true;
            }
            _ => {}
        }
    }
    if cf.font_strikethrough {
        fmt = fmt.set_font_strikethrough();
        has_format = true;
    }
    if let Some(va) = cf.font_vert_align {
        match va {
            1 => {
                fmt = fmt.set_font_script(rust_xlsxwriter::FormatScript::Superscript);
                has_format = true;
            }
            2 => {
                fmt = fmt.set_font_script(rust_xlsxwriter::FormatScript::Subscript);
                has_format = true;
            }
            _ => {}
        }
    }

    // Number format
    if let Some(ref nf) = cf.number_format {
        if nf != "General" {
            fmt = fmt.set_num_format(nf);
            has_format = true;
        }
    }

    // Alignment
    if let Some(h) = cf.align_horizontal {
        let a = match h {
            1 => rust_xlsxwriter::FormatAlign::Left,
            2 => rust_xlsxwriter::FormatAlign::Center,
            3 => rust_xlsxwriter::FormatAlign::Right,
            4 => rust_xlsxwriter::FormatAlign::Fill,
            5 => rust_xlsxwriter::FormatAlign::Justify,
            6 => rust_xlsxwriter::FormatAlign::CenterAcross,
            7 => rust_xlsxwriter::FormatAlign::Distributed,
            _ => rust_xlsxwriter::FormatAlign::General,
        };
        fmt = fmt.set_align(a);
        has_format = true;
    }
    if let Some(v) = cf.align_vertical {
        let a = match v {
            1 => rust_xlsxwriter::FormatAlign::Top,
            2 => rust_xlsxwriter::FormatAlign::VerticalCenter,
            3 => rust_xlsxwriter::FormatAlign::Bottom,
            4 => rust_xlsxwriter::FormatAlign::VerticalJustify,
            5 => rust_xlsxwriter::FormatAlign::VerticalDistributed,
            _ => rust_xlsxwriter::FormatAlign::Bottom,
        };
        fmt = fmt.set_align(a);
        has_format = true;
    }
    if cf.align_wrap_text {
        fmt = fmt.set_text_wrap();
        has_format = true;
    }
    if cf.align_shrink_to_fit {
        fmt = fmt.set_shrink();
        has_format = true;
    }
    if cf.align_indent > 0 {
        fmt = fmt.set_indent(cf.align_indent);
        has_format = true;
    }
    if cf.align_text_rotation != 0 {
        fmt = fmt.set_rotation(cf.align_text_rotation);
        has_format = true;
    }

    // Fill
    if let Some(ft) = cf.fill_type {
        fmt = fmt.set_pattern(fill_pattern_from_u8(ft));
        has_format = true;
    }
    if let Some(ref sc) = cf.fill_start_color {
        if let Some(clr) = parse_color_str(sc) {
            fmt = fmt.set_background_color(clr);
            has_format = true;
        }
    }
    if let Some(ref ec) = cf.fill_end_color {
        if let Some(clr) = parse_color_str(ec) {
            fmt = fmt.set_foreground_color(clr);
            has_format = true;
        }
    }

    // Border
    if let Some(s) = cf.border_left_style {
        fmt = fmt.set_border_left(border_style_from_u8(s));
        has_format = true;
    }
    if let Some(ref c) = cf.border_left_color {
        if let Some(clr) = parse_color_str(c) {
            fmt = fmt.set_border_left_color(clr);
            has_format = true;
        }
    }
    if let Some(s) = cf.border_right_style {
        fmt = fmt.set_border_right(border_style_from_u8(s));
        has_format = true;
    }
    if let Some(ref c) = cf.border_right_color {
        if let Some(clr) = parse_color_str(c) {
            fmt = fmt.set_border_right_color(clr);
            has_format = true;
        }
    }
    if let Some(s) = cf.border_top_style {
        fmt = fmt.set_border_top(border_style_from_u8(s));
        has_format = true;
    }
    if let Some(ref c) = cf.border_top_color {
        if let Some(clr) = parse_color_str(c) {
            fmt = fmt.set_border_top_color(clr);
            has_format = true;
        }
    }
    if let Some(s) = cf.border_bottom_style {
        fmt = fmt.set_border_bottom(border_style_from_u8(s));
        has_format = true;
    }
    if let Some(ref c) = cf.border_bottom_color {
        if let Some(clr) = parse_color_str(c) {
            fmt = fmt.set_border_bottom_color(clr);
            has_format = true;
        }
    }
    if let Some(s) = cf.border_diagonal_style {
        fmt = fmt.set_border_diagonal(border_style_from_u8(s));
        has_format = true;
    }
    if let Some(ref c) = cf.border_diagonal_color {
        if let Some(clr) = parse_color_str(c) {
            fmt = fmt.set_border_diagonal_color(clr);
            has_format = true;
        }
    }
    if cf.border_diagonal_up || cf.border_diagonal_down {
        let diag_type = match (cf.border_diagonal_up, cf.border_diagonal_down) {
            (true, true) => rust_xlsxwriter::FormatDiagonalBorder::BorderUpDown,
            (true, false) => rust_xlsxwriter::FormatDiagonalBorder::BorderUp,
            (false, true) => rust_xlsxwriter::FormatDiagonalBorder::BorderDown,
            _ => unreachable!(),
        };
        fmt = fmt.set_border_diagonal_type(diag_type);
        has_format = true;
    }

    // Protection
    if let Some(locked) = cf.protection_locked {
        if locked {
            fmt = fmt.set_locked();
        } else {
            fmt = fmt.set_unlocked();
        }
        has_format = true;
    }
    if let Some(hidden) = cf.protection_hidden {
        if hidden {
            fmt = fmt.set_hidden();
        }
        has_format = true;
    }

    (fmt, has_format)
}
