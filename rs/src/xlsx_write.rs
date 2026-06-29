//! Ghi lại file .xlsx sau khi sửa nội dung 1 sheet.
//!
//! Chiến lược: copy y nguyên TOÀN BỘ entry trong zip gốc, CHỈ thay thế nội
//! dung file XML của đúng sheet đang sửa. Điều này đảm bảo không làm hỏng
//! style, theme, ảnh, sheet khác, v.v. — tương tự việc Python chỉ
//! `ws.cell(...).value = ...` rồi `wb.save()` (openpyxl tự giữ các phần khác),
//! nhưng ở đây ta làm thủ công ở mức zip entry.
//!
//! Giá trị cell luôn được ghi dưới dạng `t="inlineStr"` (`<is><t>...</t></is>`)
//! thay vì dùng sharedStrings. Lý do: inline string không cần đồng bộ index
//! với xl/sharedStrings.xml (nguồn lỗi phổ biến nếu tính sai), Excel đọc loại
//! này hoàn toàn bình thường, chỉ đổi nhẹ cách lưu trữ nội bộ.

use std::io::{Read, Seek, Write};

use quick_xml::events::{BytesStart, BytesText, Event};
use quick_xml::writer::Writer;
use zip::write::FileOptions;
use zip::{ZipArchive, ZipWriter};

use crate::coord::col_row_to_ref;
use crate::error::AppResult;
use crate::model::{CellStyle, SheetData};

/// Escape ký tự đặc biệt XML trong text node. quick-xml's BytesText::new
/// không tự escape, nên ta escape thủ công trước khi đưa vào BytesText.
fn xml_escape_text(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    for ch in s.chars() {
        match ch {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            // Excel chấp nhận literal newline trong <t xml:space="preserve">,
            // không cần escape \n / \r ở mức XML text.
            _ => out.push(ch),
        }
    }
    out
}

/// Xác định attribute t="..." và text ghi vào <v> cho 1 cell CÓ công thức,
/// dựa trên giá trị HIỂN THỊ (đã cache) của nó:
///   - Số hợp lệ -> không cần t (mặc định numeric), v = chính số đó
///   - "TRUE"/"FALSE" -> t="b", v = "1"/"0"
///   - Còn lại (text/lỗi #REF! #DIV/0! ...) -> t="str", v = chính text đó
///     (lỗi công thức Excel chuẩn dùng t="e", nhưng dùng "str" vẫn hiển
///     thị đúng nội dung lỗi dạng text, đủ dùng cho mục đích của plugin)
fn formula_cached_type_and_value(value: &str) -> (Option<&'static str>, String) {
    if value.parse::<f64>().is_ok() {
        (None, value.to_string())
    } else if value == "TRUE" {
        (Some("b"), "1".to_string())
    } else if value == "FALSE" {
        (Some("b"), "0".to_string())
    } else {
        (Some("str"), value.to_string())
    }
}

/// Render lại 1 sheet XML hoàn toàn mới từ SheetData, dùng cấu trúc OOXML
/// tối giản nhưng hợp lệ: <worksheet><sheetData><row><c>...</c></row></sheetData>
/// <mergeCells>...</mergeCells></worksheet>
///
/// Style của từng cell ĐƯỢC giữ lại qua attribute `s="N"` (lấy từ
/// `sheet.cell_style_id`), N là index vào `<cellXfs>` của xl/styles.xml.
/// Việc đảm bảo styles.xml có đủ entry tại index N là trách nhiệm của
/// `sync_styles_xml()` (gọi riêng, trước khi ghi sheet XML vào zip).
pub fn render_sheet_xml(sheet: &SheetData) -> AppResult<Vec<u8>> {
    let mut buf = Vec::new();
    {
        let mut w = Writer::new(&mut buf);

        w.write_event(Event::Decl(quick_xml::events::BytesDecl::new(
            "1.0", Some("UTF-8"), Some("yes"),
        )))?;

        let mut root = BytesStart::new("worksheet");
        root.push_attribute((
            "xmlns",
            "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        ));
        w.write_event(Event::Start(root))?;

        // <dimension ref="A1:F20"/>
        let dim_ref = format!(
            "A1:{}",
            col_row_to_ref(sheet.max_col.max(1), sheet.max_row.max(1))
        );
        let mut dim = BytesStart::new("dimension");
        dim.push_attribute(("ref", dim_ref.as_str()));
        w.write_event(Event::Empty(dim))?;

        // <sheetData>
        w.write_event(Event::Start(BytesStart::new("sheetData")))?;

        let max_row = sheet.max_row.max(1);
        let max_col = sheet.max_col.max(1);

        for row in 1..=max_row {
            let mut row_tag = BytesStart::new("row");
            let row_str = row.to_string();
            row_tag.push_attribute(("r", row_str.as_str()));
            w.write_event(Event::Start(row_tag))?;

            for col in 1..=max_col {
                let cell_ref = col_row_to_ref(col, row);
                let style_id = sheet.cell_style_id.get(&(row, col)).copied();
                match sheet.get(row, col) {
                    None => {
                        // Cell rỗng nhưng có style riêng (ví dụ chỉ tô màu nền,
                        // chưa nhập text) -> vẫn cần ghi <c> để giữ style.
                        if let Some(sid) = style_id {
                            if sid != 0 {
                                let mut c_tag = BytesStart::new("c");
                                c_tag.push_attribute(("r", cell_ref.as_str()));
                                let sid_str = sid.to_string();
                                c_tag.push_attribute(("s", sid_str.as_str()));
                                w.write_event(Event::Empty(c_tag))?;
                            }
                        }
                    }
                    Some(value) => {
                        let mut c_tag = BytesStart::new("c");
                        c_tag.push_attribute(("r", cell_ref.as_str()));
                        let sid_str = style_id.unwrap_or(0).to_string();
                        if style_id.unwrap_or(0) != 0 {
                            c_tag.push_attribute(("s", sid_str.as_str()));
                        }

                        if let Some(formula) = sheet.formulas.get(&(row, col)) {
                            // Cell có công thức: ghi <f>...</f><v>...</v>,
                            // KHÔNG dùng inlineStr — Excel cần <v> đúng
                            // kiểu (number/string/bool) để hiển thị ngay
                            // khi mở file, trước khi tự tính lại công thức.
                            let (t_attr, v_text) = formula_cached_type_and_value(value);
                            if let Some(t) = t_attr {
                                c_tag.push_attribute(("t", t));
                            }
                            w.write_event(Event::Start(c_tag))?;

                            w.write_event(Event::Start(BytesStart::new("f")))?;
                            w.write_event(Event::Text(BytesText::from_escaped(xml_escape_text(
                                formula,
                            ))))?;
                            w.write_event(Event::End(quick_xml::events::BytesEnd::new("f")))?;

                            w.write_event(Event::Start(BytesStart::new("v")))?;
                            w.write_event(Event::Text(BytesText::from_escaped(xml_escape_text(
                                &v_text,
                            ))))?;
                            w.write_event(Event::End(quick_xml::events::BytesEnd::new("v")))?;

                            w.write_event(Event::End(quick_xml::events::BytesEnd::new("c")))?;
                        } else {
                            c_tag.push_attribute(("t", "inlineStr"));
                            w.write_event(Event::Start(c_tag))?;

                            w.write_event(Event::Start(BytesStart::new("is")))?;

                            let mut t_tag = BytesStart::new("t");
                            // Giữ khoảng trắng đầu/cuối nguyên vẹn nếu có.
                            t_tag.push_attribute(("xml:space", "preserve"));
                            w.write_event(Event::Start(t_tag))?;
                            w.write_event(Event::Text(BytesText::from_escaped(xml_escape_text(
                                value,
                            ))))?;
                            w.write_event(Event::End(quick_xml::events::BytesEnd::new("t")))?;

                            w.write_event(Event::End(quick_xml::events::BytesEnd::new("is")))?;
                            w.write_event(Event::End(quick_xml::events::BytesEnd::new("c")))?;
                        }
                    }
                }
            }

            w.write_event(Event::End(quick_xml::events::BytesEnd::new("row")))?;
        }

        w.write_event(Event::End(quick_xml::events::BytesEnd::new("sheetData")))?;

        // <mergeCells count="N">...</mergeCells>
        if !sheet.merges.is_empty() {
            let mut mc = BytesStart::new("mergeCells");
            let count_str = sheet.merges.len().to_string();
            mc.push_attribute(("count", count_str.as_str()));
            w.write_event(Event::Start(mc))?;

            for m in &sheet.merges {
                let range = format!(
                    "{}:{}",
                    col_row_to_ref(m.min_col, m.min_row),
                    col_row_to_ref(m.max_col, m.max_row)
                );
                let mut mc_cell = BytesStart::new("mergeCell");
                mc_cell.push_attribute(("ref", range.as_str()));
                w.write_event(Event::Empty(mc_cell))?;
            }

            w.write_event(Event::End(quick_xml::events::BytesEnd::new(
                "mergeCells",
            )))?;
        }

        w.write_event(Event::End(quick_xml::events::BytesEnd::new("worksheet")))?;
    }
    Ok(buf)
}

/// Ghi lại toàn bộ file .xlsx vào bộ nhớ: copy mọi entry gốc, chỉ thay thế
/// nội dung của các entry có trong `replacements`. Trả về bytes hoàn chỉnh
/// của file .xlsx mới — caller chịu trách nhiệm ghi ra đĩa (atomic).
pub fn write_xlsx_replacing_entries<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    replacements: &[(&str, Vec<u8>)],
) -> AppResult<Vec<u8>> {
    let mut out_buf = Vec::new();
    {
        let cursor = std::io::Cursor::new(&mut out_buf);
        let mut zw = ZipWriter::new(cursor);
        let options = FileOptions::default()
            .compression_method(zip::CompressionMethod::Deflated);

        let mut written: std::collections::HashSet<&str> = std::collections::HashSet::new();

        for i in 0..archive.len() {
            let mut entry = archive.by_index(i)?;
            let name = entry.name().to_string();

            zw.start_file(&name, options)?;
            if let Some((path, bytes)) = replacements.iter().find(|(p, _)| **p == name) {
                zw.write_all(bytes)?;
                written.insert(path);
            } else {
                let mut buf = Vec::with_capacity(entry.size() as usize);
                entry.read_to_end(&mut buf)?;
                zw.write_all(&buf)?;
            }
        }

        // Nếu 1 path trong replacements không tồn tại sẵn trong zip gốc
        // (ví dụ file không có sharedStrings.xml/styles.xml) -> thêm mới.
        for (path, bytes) in replacements {
            if !written.contains(*path) {
                zw.start_file(*path, options)?;
                zw.write_all(bytes)?;
            }
        }

        zw.finish()?;
    }
    Ok(out_buf)
}

/// Đồng bộ lại xl/styles.xml gốc để chứa đủ font/fill/cellXfs cho TOÀN BỘ
/// style trong `sheet.styles` (bao gồm style mới do user thêm qua lệnh
/// setbg, vốn không tồn tại trong file gốc). Style cũ (index < số cellXfs
/// gốc) được giữ NGUYÊN VẸN bytes-for-bytes — chỉ append thêm phần tử mới
/// vào cuối <fonts>/<fills>/<cellXfs>, nên không ảnh hưởng cell nào khác.
///
/// Trả về `None` nếu không có style mới nào cần thêm (tức styles.xml gốc đã
/// đủ dùng) — caller có thể bỏ qua, không cần ghi lại styles.xml.
pub fn sync_styles_xml(
    original_styles_xml: Option<&[u8]>,
    all_styles: &[CellStyle],
) -> AppResult<Option<Vec<u8>>> {
    // Số cellXfs gốc đã tồn tại trong file (= độ dài Vec<CellStyle> mà
    // read_styles() trả về khi đọc lần đầu). Nếu all_styles không có gì
    // mới hơn con số đó thì không cần sửa styles.xml.
    let existing_xml = match original_styles_xml {
        Some(b) => b.to_vec(),
        None => default_styles_xml(),
    };

    let existing_count = count_cell_xfs(&existing_xml)?;
    if all_styles.len() <= existing_count {
        return Ok(None);
    }

    let new_styles = &all_styles[existing_count..];

    let mut reader = quick_xml::reader::Reader::from_reader(existing_xml.as_slice());
    reader.trim_text(false);
    let mut writer = Writer::new(Vec::new());
    let mut buf = Vec::new();

    // Đếm sẵn số font/fill hiện có để biết fontId/fillId mới sẽ bắt đầu từ đâu.
    let existing_font_count = count_tag(&existing_xml, b"fonts", b"font")?;
    let existing_fill_count = count_tag(&existing_xml, b"fills", b"fill")?;

    let mut next_font_id = existing_font_count as u32;
    let mut next_fill_id = existing_fill_count as u32;
    // (font_id, fill_id) cho mỗi style mới, theo đúng thứ tự sẽ append vào cellXfs.
    let mut new_xf_ids: Vec<(u32, u32)> = Vec::with_capacity(new_styles.len());
    let mut new_font_tags: Vec<String> = Vec::new();
    let mut new_fill_tags: Vec<String> = Vec::new();

    for style in new_styles {
        let needs_font = style.bold || style.italic || style.font_color.is_some();
        let font_id = if needs_font {
            let tag = format!(
                "<font>{}{}{}</font>",
                if style.bold { "<b/>" } else { "" },
                if style.italic { "<i/>" } else { "" },
                style
                    .font_color
                    .map(|c| format!("<color rgb=\"FF{}\"/>", c.to_hex()))
                    .unwrap_or_default(),
            );
            new_font_tags.push(tag);
            let id = next_font_id;
            next_font_id += 1;
            id
        } else {
            0 // font mặc định (id 0 luôn tồn tại theo chuẩn OOXML)
        };

        let fill_id = if let Some(bg) = style.bg_color {
            let tag = format!(
                "<fill><patternFill patternType=\"solid\"><fgColor rgb=\"FF{hex}\"/><bgColor rgb=\"FF{hex}\"/></patternFill></fill>",
                hex = bg.to_hex()
            );
            new_fill_tags.push(tag);
            let id = next_fill_id;
            next_fill_id += 1;
            id
        } else {
            0 // "no fill" (id 0 theo chuẩn OOXML mặc định)
        };

        new_xf_ids.push((font_id, fill_id));
    }

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Eof => break,
            Event::End(e) if e.local_name().as_ref() == b"fonts" => {
                for tag in &new_font_tags {
                    writer.get_mut().extend_from_slice(tag.as_bytes());
                }
                writer.write_event(Event::End(e))?;
            }
            Event::End(e) if e.local_name().as_ref() == b"fills" => {
                for tag in &new_fill_tags {
                    writer.get_mut().extend_from_slice(tag.as_bytes());
                }
                writer.write_event(Event::End(e))?;
            }
            Event::End(e) if e.local_name().as_ref() == b"cellXfs" => {
                for (font_id, fill_id) in &new_xf_ids {
                    let tag = format!(
                        "<xf numFmtId=\"0\" fontId=\"{font_id}\" fillId=\"{fill_id}\" borderId=\"0\" xfId=\"0\"/>"
                    );
                    writer.get_mut().extend_from_slice(tag.as_bytes());
                }
                writer.write_event(Event::End(e))?;
            }
            // Cập nhật lại count="N" trên <fonts>, <fills>, <cellXfs> để khớp
            // số lượng thực tế sau khi append (Excel không bắt buộc nhưng
            // 1 số phiên bản LibreOffice kiểm tra, nên cập nhật cho chuẩn).
            Event::Start(e) if e.local_name().as_ref() == b"fonts" => {
                writer.write_event(Event::Start(rewrite_count(&e, existing_font_count + new_font_tags.len())?))?;
            }
            Event::Start(e) if e.local_name().as_ref() == b"fills" => {
                writer.write_event(Event::Start(rewrite_count(&e, existing_fill_count + new_fill_tags.len())?))?;
            }
            Event::Start(e) if e.local_name().as_ref() == b"cellXfs" => {
                writer.write_event(Event::Start(rewrite_count(&e, existing_count + new_xf_ids.len())?))?;
            }
            other => {
                writer.write_event(other)?;
            }
        }
        buf.clear();
    }

    Ok(Some(writer.into_inner()))
}

/// Đếm số tag con trực tiếp (ví dụ <font> trong <fonts>) để biết id tiếp
/// theo sẽ append vào đâu.
fn count_tag(xml: &[u8], parent_local: &[u8], child_local: &[u8]) -> AppResult<usize> {
    let mut reader = quick_xml::reader::Reader::from_reader(xml);
    reader.trim_text(true);
    let mut buf = Vec::new();
    let mut in_parent = false;
    let mut depth = 0i32;
    let mut count = 0usize;
    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) if e.local_name().as_ref() == parent_local => {
                in_parent = true;
                depth = 0;
            }
            Event::End(e) if e.local_name().as_ref() == parent_local => {
                in_parent = false;
            }
            Event::Start(e) if in_parent && e.local_name().as_ref() == child_local && depth == 0 => {
                count += 1;
                depth += 1;
            }
            Event::Start(_) if in_parent => depth += 1,
            Event::End(_) if in_parent && depth > 0 => depth -= 1,
            Event::Empty(e) if in_parent && e.local_name().as_ref() == child_local && depth == 0 => {
                count += 1;
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(count)
}

fn count_cell_xfs(xml: &[u8]) -> AppResult<usize> {
    count_tag(xml, b"cellXfs", b"xf")
}

/// Tạo lại 1 BytesStart với attribute count="N" được ghi đè (hoặc thêm mới
/// nếu chưa có), giữ nguyên mọi attribute khác.
fn rewrite_count<'a>(e: &BytesStart<'a>, new_count: usize) -> AppResult<BytesStart<'static>> {
    let mut new_tag = BytesStart::new(String::from_utf8(e.name().as_ref().to_vec())?);
    let mut had_count = false;
    for attr in e.attributes() {
        let attr = attr?;
        if attr.key.as_ref() == b"count" {
            let count_str = new_count.to_string();
            new_tag.push_attribute(("count", count_str.as_str()));
            had_count = true;
        } else {
            let key = String::from_utf8(attr.key.as_ref().to_vec())?;
            let val = attr.unescape_value()?.into_owned();
            new_tag.push_attribute((key.as_str(), val.as_str()));
        }
    }
    if !had_count {
        let count_str = new_count.to_string();
        new_tag.push_attribute(("count", count_str.as_str()));
    }
    Ok(new_tag)
}

/// styles.xml tối giản dùng khi file gốc KHÔNG có xl/styles.xml (trường hợp
/// hiếm, vì openpyxl luôn ghi styles.xml, nhưng vẫn cần fallback an toàn).
fn default_styles_xml() -> Vec<u8> {
    br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
<fills count="1"><fill><patternFill patternType="none"/></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>"#
        .to_vec()
}



/// Đổi tên 1 sheet trực tiếp trong xl/workbook.xml (chỉ đổi attribute name=
/// của đúng <sheet> có r:id tương ứng), giữ nguyên mọi thứ khác.
pub fn rename_sheet_in_workbook_xml(xml: &[u8], target_rid: &str, new_name: &str) -> AppResult<Vec<u8>> {
    let mut reader = quick_xml::reader::Reader::from_reader(xml);
    reader.trim_text(false);
    let mut writer = Writer::new(Vec::new());
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Eof => break,
            Event::Empty(e) if e.local_name().as_ref() == b"sheet" => {
                let is_target = e.attributes().filter_map(|a| a.ok()).any(|a| {
                    (a.key.as_ref() == b"r:id" || a.key.as_ref().ends_with(b":id"))
                        && a.unescape_value().map(|v| v == target_rid).unwrap_or(false)
                });
                if is_target {
                    let mut new_tag = BytesStart::new("sheet");
                    for a in e.attributes().filter_map(|a| a.ok()) {
                        let key = String::from_utf8_lossy(a.key.as_ref()).into_owned();
                        if key == "name" {
                            new_tag.push_attribute(("name", new_name));
                        } else {
                            let val = a.unescape_value()?.into_owned();
                            new_tag.push_attribute((key.as_str(), val.as_str()));
                        }
                    }
                    writer.write_event(Event::Empty(new_tag))?;
                } else {
                    writer.write_event(Event::Empty(e))?;
                }
            }
            other => {
                writer.write_event(other)?;
            }
        }
        buf.clear();
    }

    Ok(writer.into_inner())
}
