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
use crate::model::SheetData;

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

/// Render lại 1 sheet XML hoàn toàn mới từ SheetData, dùng cấu trúc OOXML
/// tối giản nhưng hợp lệ: <worksheet><sheetData><row><c>...</c></row></sheetData>
/// <mergeCells>...</mergeCells></worksheet>
///
/// Đây KHÔNG cố giữ style/format gốc của từng cell (đó nằm trong styles.xml
/// qua attribute s="..." trên <c>, mà ta không parse ở bước đọc). Việc này
/// khớp với hành vi excel.py gốc: hiển thị/sửa dạng bảng ASCII thuần text,
/// không can thiệp định dạng ô.
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
                match sheet.get(row, col) {
                    None => {
                        // Cell rỗng: bỏ qua hoàn toàn (Excel coi cell không
                        // xuất hiện trong XML là rỗng) -> giữ file nhỏ gọn.
                    }
                    Some(value) => {
                        let mut c_tag = BytesStart::new("c");
                        c_tag.push_attribute(("r", cell_ref.as_str()));
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
/// nội dung của `target_sheet_path` bằng `new_sheet_xml`. Trả về bytes hoàn
/// chỉnh của file .xlsx mới — caller chịu trách nhiệm ghi ra đĩa (atomic).
pub fn write_xlsx_replacing_sheet<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    target_sheet_path: &str,
    new_sheet_xml: &[u8],
) -> AppResult<Vec<u8>> {
    let mut out_buf = Vec::new();
    {
        let cursor = std::io::Cursor::new(&mut out_buf);
        let mut zw = ZipWriter::new(cursor);
        let options = FileOptions::default()
            .compression_method(zip::CompressionMethod::Deflated);

        for i in 0..archive.len() {
            let mut entry = archive.by_index(i)?;
            let name = entry.name().to_string();

            zw.start_file(&name, options)?;
            if name == target_sheet_path {
                zw.write_all(new_sheet_xml)?;
            } else {
                let mut buf = Vec::with_capacity(entry.size() as usize);
                entry.read_to_end(&mut buf)?;
                zw.write_all(&buf)?;
            }
        }

        zw.finish()?;
    }
    Ok(out_buf)
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
