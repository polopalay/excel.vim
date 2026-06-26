//! Đọc cấu trúc workbook từ file .xlsx (là 1 file ZIP chứa các XML con).
//!
//! Tương đương phần "đọc" của openpyxl.load_workbook(), nhưng viết tay bằng
//! quick-xml để tránh phụ thuộc 1 crate xlsx ngoài có API có thể thay đổi.

use std::collections::BTreeMap;
use std::io::Read;

use quick_xml::events::Event;
use quick_xml::reader::Reader;
use zip::ZipArchive;

use crate::coord::ref_to_col_row;
use crate::error::{AppError, AppResult};
use crate::model::{CellStyle, MergeRange, Rgb, SheetData, SheetEntry};

/// Đọc toàn bộ bytes của 1 entry trong zip theo path, trả lỗi rõ ràng nếu
/// không tồn tại (file .xlsx bị hỏng / không đúng cấu trúc OOXML).
pub fn read_zip_entry<R: Read + std::io::Seek>(
    archive: &mut ZipArchive<R>,
    path: &str,
) -> AppResult<Vec<u8>> {
    let mut file = archive
        .by_name(path)
        .map_err(|_| AppError(format!("Missing entry in xlsx: {path}")))?;
    let mut buf = Vec::new();
    file.read_to_end(&mut buf)?;
    Ok(buf)
}

/// Lấy giá trị 1 attribute trong tag XML hiện tại, nếu có.
fn get_attr(e: &quick_xml::events::BytesStart, key: &str) -> AppResult<Option<String>> {
    for attr in e.attributes() {
        let attr = attr?;
        if attr.key.as_ref() == key.as_bytes() {
            return Ok(Some(attr.unescape_value()?.into_owned()));
        }
    }
    Ok(None)
}

/// Đọc xl/sharedStrings.xml -> Vec<String> (index 0-based khớp với attribute
/// `t="s"` trong sheet XML, giống cách Excel/openpyxl dùng).
///
/// Mỗi `<si>` có thể chứa 1 `<t>` đơn giản, hoặc nhiều `<r><t>...</t></r>`
/// (rich text runs) -> ta nối toàn bộ text con lại, đủ dùng cho mục đích
/// hiển thị dạng bảng ASCII (không cần giữ định dạng rich text).
pub fn read_shared_strings(bytes: Option<&[u8]>) -> AppResult<Vec<String>> {
    let Some(bytes) = bytes else {
        return Ok(Vec::new());
    };
    let mut reader = Reader::from_reader(bytes);
    reader.trim_text(false);

    let mut strings = Vec::new();
    let mut current = String::new();
    let mut in_si = false;
    let mut in_t = false;
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) if e.local_name().as_ref() == b"si" => {
                in_si = true;
                current.clear();
            }
            Event::End(e) if e.local_name().as_ref() == b"si" => {
                in_si = false;
                strings.push(std::mem::take(&mut current));
            }
            Event::Start(e) if in_si && e.local_name().as_ref() == b"t" => {
                in_t = true;
            }
            Event::End(e) if e.local_name().as_ref() == b"t" => {
                in_t = false;
            }
            Event::Text(t) if in_si && in_t => {
                current.push_str(&t.unescape()?);
            }
            // <si><t/></si> rỗng (self-closing) -> không có Text event, ok giữ "".
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    Ok(strings)
}

/// Đọc xl/styles.xml -> Vec<CellStyle>, đã resolve sẵn theo đúng INDEX của
/// `<cellXfs>` (chính là giá trị attribute `s="N"` trên mỗi `<c>` trong sheet
/// XML). Tương đương việc openpyxl tự resolve `cell.font` / `cell.fill` cho
/// từng cell, nhưng ở đây ta resolve 1 lần cho toàn bộ bảng cellXfs.
///
/// Cấu trúc styles.xml liên quan:
///   <fonts><font>...<b/><i/><color rgb="FFRRGGBB"/></font>...</fonts>
///   <fills><fill><patternFill patternType="solid"><fgColor rgb="FFRRGGBB"/></patternFill></fill>...</fills>
///   <cellXfs><xf fontId="N" fillId="M" .../>...</cellXfs>
/// `s="N"` trên <c> chính là index vào <cellXfs>, N trỏ tới fontId/fillId.
pub fn read_styles(bytes: Option<&[u8]>) -> AppResult<Vec<CellStyle>> {
    let Some(bytes) = bytes else {
        // Không có styles.xml -> coi như mọi cell dùng style mặc định (id 0).
        return Ok(vec![CellStyle::default()]);
    };

    // --- 1. Đọc <fonts> -> Vec<(bold, italic, color)> theo fontId ---
    #[derive(Default, Clone)]
    struct FontInfo {
        bold: bool,
        italic: bool,
        color: Option<Rgb>,
    }
    #[derive(Default, Clone)]
    struct FillInfo {
        bg: Option<Rgb>,
    }

    let mut fonts: Vec<FontInfo> = Vec::new();
    let mut fills: Vec<FillInfo> = Vec::new();
    let mut xfs: Vec<(u32, u32)> = Vec::new(); // (font_id, fill_id) theo cellXfs index

    let mut reader = Reader::from_reader(bytes);
    reader.trim_text(true);
    let mut buf = Vec::new();

    let mut in_fonts = false;
    let mut in_fills = false;
    let mut in_cell_xfs = false;
    let mut in_pattern_fill = false;
    let mut cur_font: Option<FontInfo> = None;
    let mut cur_fill: Option<FillInfo> = None;

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) if e.local_name().as_ref() == b"fonts" => in_fonts = true,
            Event::End(e) if e.local_name().as_ref() == b"fonts" => in_fonts = false,
            Event::Start(e) if e.local_name().as_ref() == b"fills" => in_fills = true,
            Event::End(e) if e.local_name().as_ref() == b"fills" => in_fills = false,
            Event::Start(e) if e.local_name().as_ref() == b"cellXfs" => in_cell_xfs = true,
            Event::End(e) if e.local_name().as_ref() == b"cellXfs" => in_cell_xfs = false,

            Event::Start(e) if in_fonts && e.local_name().as_ref() == b"font" => {
                cur_font = Some(FontInfo::default());
            }
            Event::End(e) if in_fonts && e.local_name().as_ref() == b"font" => {
                fonts.push(cur_font.take().unwrap_or_default());
            }
            Event::Start(e) | Event::Empty(e) if cur_font.is_some() && e.local_name().as_ref() == b"b" => {
                // <b/> hoặc <b val="1"/>; chỉ val="0" mới là false
                let is_false = get_attr(&e, "val")?.map(|v| v == "0").unwrap_or(false);
                if let Some(f) = cur_font.as_mut() {
                    f.bold = !is_false;
                }
            }
            Event::Start(e) | Event::Empty(e) if cur_font.is_some() && e.local_name().as_ref() == b"i" => {
                let is_false = get_attr(&e, "val")?.map(|v| v == "0").unwrap_or(false);
                if let Some(f) = cur_font.as_mut() {
                    f.italic = !is_false;
                }
            }
            Event::Start(e) | Event::Empty(e) if cur_font.is_some() && e.local_name().as_ref() == b"color" => {
                if let Some(rgb_attr) = get_attr(&e, "rgb")? {
                    if let Some(f) = cur_font.as_mut() {
                        f.color = Rgb::from_hex(&rgb_attr);
                    }
                }
                // color theme="N" (màu theo theme) — bỏ qua, coi như màu mặc định,
                // vì ta không parse theme1.xml (đủ dùng cho mục đích hiển thị ASCII).
            }

            Event::Start(e) if in_fills && e.local_name().as_ref() == b"fill" => {
                cur_fill = Some(FillInfo::default());
            }
            Event::End(e) if in_fills && e.local_name().as_ref() == b"fill" => {
                fills.push(cur_fill.take().unwrap_or_default());
            }
            Event::Start(e) | Event::Empty(e)
                if cur_fill.is_some() && e.local_name().as_ref() == b"patternFill" =>
            {
                let pattern_type = get_attr(&e, "patternType")?;
                in_pattern_fill = pattern_type.as_deref() == Some("solid");
            }
            Event::Start(e) | Event::Empty(e)
                if in_pattern_fill && cur_fill.is_some() && e.local_name().as_ref() == b"fgColor" =>
            {
                if let Some(rgb_attr) = get_attr(&e, "rgb")? {
                    if let Some(fl) = cur_fill.as_mut() {
                        fl.bg = Rgb::from_hex(&rgb_attr);
                    }
                }
            }

            Event::Start(e) | Event::Empty(e) if in_cell_xfs && e.local_name().as_ref() == b"xf" => {
                let font_id: u32 = get_attr(&e, "fontId")?.and_then(|s| s.parse().ok()).unwrap_or(0);
                let fill_id: u32 = get_attr(&e, "fillId")?.and_then(|s| s.parse().ok()).unwrap_or(0);
                xfs.push((font_id, fill_id));
            }

            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    if xfs.is_empty() {
        return Ok(vec![CellStyle::default()]);
    }

    let styles = xfs
        .into_iter()
        .map(|(font_id, fill_id)| {
            let font = fonts.get(font_id as usize).cloned().unwrap_or_default();
            let fill = fills.get(fill_id as usize).cloned().unwrap_or_default();
            CellStyle {
                bold: font.bold,
                italic: font.italic,
                font_color: font.color,
                bg_color: fill.bg,
            }
        })
        .collect();

    Ok(styles)
}


/// theo đúng thứ tự hiển thị trong Excel (thứ tự tag <sheet> trong workbook.xml,
/// KHÔNG phải thứ tự file sheetN.xml vật lý trong zip).
pub fn read_sheet_entries(
    workbook_xml: &[u8],
    workbook_rels_xml: &[u8],
) -> AppResult<Vec<SheetEntry>> {
    // 1. Đọc rels: rId -> target path (relative tới xl/)
    let mut rels: BTreeMap<String, String> = BTreeMap::new();
    {
        let mut reader = Reader::from_reader(workbook_rels_xml);
        reader.trim_text(true);
        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf)? {
                Event::Empty(e) | Event::Start(e) if e.local_name().as_ref() == b"Relationship" => {
                    let id = get_attr(&e, "Id")?;
                    let target = get_attr(&e, "Target")?;
                    if let (Some(id), Some(target)) = (id, target) {
                        // Target thường dạng "worksheets/sheet1.xml" (relative tới xl/)
                        let full = if target.starts_with("/xl/") {
                            target.trim_start_matches('/').to_string()
                        } else if target.starts_with("xl/") {
                            target
                        } else {
                            format!("xl/{target}")
                        };
                        rels.insert(id, full);
                    }
                }
                Event::Eof => break,
                _ => {}
            }
            buf.clear();
        }
    }

    // 2. Đọc workbook.xml: thứ tự + tên + r:id + sheetId của từng <sheet>
    let mut entries = Vec::new();
    {
        let mut reader = Reader::from_reader(workbook_xml);
        reader.trim_text(true);
        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf)? {
                Event::Empty(e) | Event::Start(e) if e.local_name().as_ref() == b"sheet" => {
                    let name = get_attr(&e, "name")?.unwrap_or_default();
                    let sheet_id: u32 = get_attr(&e, "sheetId")?
                        .and_then(|s| s.parse().ok())
                        .unwrap_or(0);
                    // r:id có namespace prefix "r:" -> tìm theo suffix vì
                    // quick-xml trả raw key kèm prefix khi không resolve namespace.
                    let mut rid = None;
                    for attr in e.attributes() {
                        let attr = attr?;
                        let key = attr.key.as_ref();
                        if key == b"r:id" || key.ends_with(b":id") {
                            rid = Some(attr.unescape_value()?.into_owned());
                        }
                    }
                    let rid = rid.ok_or_else(|| {
                        AppError(format!("Sheet '{name}' missing r:id in workbook.xml"))
                    })?;
                    let path = rels.get(&rid).cloned().ok_or_else(|| {
                        AppError(format!("Sheet '{name}' r:id={rid} not found in rels"))
                    })?;
                    entries.push(SheetEntry {
                        name,
                        path,
                        rid,
                        sheet_id,
                    });
                }
                Event::Eof => break,
                _ => {}
            }
            buf.clear();
        }
    }

    if entries.is_empty() {
        return Err(AppError("No sheets found in workbook.xml".to_string()));
    }

    Ok(entries)
}

/// Parse 1 file xl/worksheets/sheetN.xml -> SheetData (giá trị cell đã resolve
/// qua sharedStrings, danh sách merge ranges, kích thước max_row/max_col).
///
/// Tương đương việc openpyxl trả về `ws` với `data_only=True` (lấy giá trị đã
/// tính sẵn trong file, không tính lại công thức).
pub fn parse_sheet_xml(xml: &[u8], shared_strings: &[String], styles: &[CellStyle]) -> AppResult<SheetData> {
    let mut reader = Reader::from_reader(xml);
    reader.trim_text(false);
    let mut buf = Vec::new();

    let mut sheet = SheetData::default();
    sheet.styles = styles.to_vec();

    // Trạng thái đang đọc 1 <c> cell
    let mut cur_ref: Option<String> = None;
    let mut cur_type: Option<String> = None; // attribute t="s"/"str"/"b"/"inlineStr"/...
    let mut cur_style_id: Option<u32> = None; // attribute s="N"
    let mut cur_value = String::new();
    let mut in_v = false;
    let mut in_is_t = false; // trong <is><t>...</t></is> (inline string)
    let mut in_c = false;

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) | Event::Empty(e) if e.local_name().as_ref() == b"dimension" => {
                if let Some(ref_attr) = get_attr(&e, "ref")? {
                    // ref dạng "A1:F20" hoặc chỉ "A1" nếu sheet trống/1 cell
                    if let Some((_, _, max_c, max_r)) =
                        crate::coord::parse_range(&ref_attr).ok()
                    {
                        sheet.max_row = max_r;
                        sheet.max_col = max_c;
                    }
                }
            }
            Event::Start(e) if e.local_name().as_ref() == b"c" => {
                in_c = true;
                cur_ref = get_attr(&e, "r")?;
                cur_type = get_attr(&e, "t")?;
                cur_style_id = get_attr(&e, "s")?.and_then(|s| s.parse().ok());
                cur_value.clear();
            }
            Event::Empty(e) if e.local_name().as_ref() == b"c" => {
                // <c r="A1"/> hoàn toàn rỗng (tự đóng, không có <v> con).
                // Không có giá trị để lưu, nhưng vẫn cần mở rộng bounds để
                // dimension/max_row/max_col phản ánh đúng kích thước sheet.
                if let Some(r) = get_attr(&e, "r")? {
                    if let Ok((col, row)) = ref_to_col_row(&r) {
                        if row > sheet.max_row {
                            sheet.max_row = row;
                        }
                        if col > sheet.max_col {
                            sheet.max_col = col;
                        }
                        if let Some(sid) = get_attr(&e, "s")?.and_then(|s| s.parse::<u32>().ok()) {
                            if sid != 0 {
                                sheet.cell_style_id.insert((row, col), sid);
                            }
                        }
                    }
                }
            }
            Event::End(e) if e.local_name().as_ref() == b"c" => {
                in_c = false;
                if let Some(cell_ref) = cur_ref.take() {
                    let (col, row) = ref_to_col_row(&cell_ref)?;
                    if let Some(sid) = cur_style_id {
                        if sid != 0 {
                            sheet.cell_style_id.insert((row, col), sid);
                        }
                    }
                    let resolved = resolve_value(&cur_type, &cur_value, shared_strings);
                    if !resolved.is_empty() {
                        sheet.set(row, col, resolved);
                    } else {
                        // Vẫn cập nhật bounds dù cell rỗng, để giữ đúng max_row/max_col
                        if row > sheet.max_row {
                            sheet.max_row = row;
                        }
                        if col > sheet.max_col {
                            sheet.max_col = col;
                        }
                    }
                }
                cur_type = None;
                cur_style_id = None;
                cur_value.clear();
            }
            Event::Start(e) if in_c && e.local_name().as_ref() == b"v" => {
                in_v = true;
            }
            Event::End(e) if e.local_name().as_ref() == b"v" => {
                in_v = false;
            }
            Event::Start(e) if in_c && e.local_name().as_ref() == b"is" => {
                // inline string container <is><t>...</t></is>
                let _ = e;
            }
            Event::Start(e) if in_c && e.local_name().as_ref() == b"t" => {
                in_is_t = true;
            }
            Event::End(e) if e.local_name().as_ref() == b"t" => {
                in_is_t = false;
            }
            Event::Text(t) if in_v || in_is_t => {
                cur_value.push_str(&t.unescape()?);
            }
            Event::Start(e) | Event::Empty(e) if e.local_name().as_ref() == b"mergeCell" => {
                if let Some(range) = get_attr(&e, "ref")? {
                    let (min_col, min_row, max_col, max_row) = crate::coord::parse_range(&range)?;
                    sheet.merges.push(MergeRange {
                        min_row,
                        max_row,
                        min_col,
                        max_col,
                    });
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    sheet.recompute_bounds();
    Ok(sheet)
}

/// Resolve giá trị thô (<v> text) theo type attribute, giống cách Excel diễn
/// giải nội dung cell:
///   t="s"        -> index vào sharedStrings
///   t="str"      -> kết quả công thức dạng chuỗi, dùng trực tiếp
///   t="inlineStr"-> đã resolve sẵn trong cur_value qua <is><t>
///   t="b"        -> boolean "1"/"0" -> "TRUE"/"FALSE" (giống openpyxl)
///   None / "n"   -> số, dùng trực tiếp (giữ nguyên text số trong XML)
fn resolve_value(cell_type: &Option<String>, raw: &str, shared: &[String]) -> String {
    match cell_type.as_deref() {
        Some("s") => raw
            .parse::<usize>()
            .ok()
            .and_then(|i| shared.get(i))
            .cloned()
            .unwrap_or_default(),
        Some("b") => {
            if raw == "1" {
                "TRUE".to_string()
            } else {
                "FALSE".to_string()
            }
        }
        _ => raw.to_string(),
    }
}
