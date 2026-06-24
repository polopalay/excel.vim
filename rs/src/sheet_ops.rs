//! Quản lý sheet (list/create/rename/delete) — port từ
//! `list_sheets/create_sheet/rename_sheet/delete_sheet` trong excel.py,
//! thao tác trực tiếp trên xl/workbook.xml + xl/workbook.xml.rels +
//! [Content_Types].xml, copy mọi entry khác y nguyên.

use std::io::{Read, Seek, Write};

use quick_xml::events::{BytesStart, Event};
use quick_xml::writer::Writer;
use zip::write::FileOptions;
use zip::{ZipArchive, ZipWriter};

use crate::bail;
use crate::error::AppResult;
use crate::model::SheetEntry;
use crate::xlsx_read::read_zip_entry;

pub fn list_sheet_names(entries: &[SheetEntry]) -> Vec<String> {
    entries.iter().map(|e| e.name.clone()).collect()
}

/// Tìm sheetId lớn nhất hiện có để sinh sheetId mới (giống cách Excel/openpyxl
/// đảm bảo sheetId không bao giờ trùng, kể cả sau khi xoá sheet khác).
fn next_sheet_id(entries: &[SheetEntry]) -> u32 {
    entries.iter().map(|e| e.sheet_id).max().unwrap_or(0) + 1
}

/// Tìm số thứ tự sheetN.xml tiếp theo còn trống trong xl/worksheets/, để file
/// mới không đè lên sheet đã tồn tại.
fn next_sheet_file_index<R: Read + Seek>(archive: &mut ZipArchive<R>) -> u32 {
    let mut max_n = 0u32;
    for i in 0..archive.len() {
        if let Ok(entry) = archive.by_index(i) {
            let name = entry.name();
            if let Some(rest) = name
                .strip_prefix("xl/worksheets/sheet")
                .and_then(|s| s.strip_suffix(".xml"))
            {
                if let Ok(n) = rest.parse::<u32>() {
                    max_n = max_n.max(n);
                }
            }
        }
    }
    max_n + 1
}

/// Tạo sheet mới: thêm 1 entry sheetN.xml rỗng + cập nhật workbook.xml,
/// workbook.xml.rels, [Content_Types].xml. Trả về bytes file .xlsx mới.
pub fn create_sheet<R: Read + Seek>(
    mut archive: ZipArchive<R>,
    entries: &[SheetEntry],
    new_name: &str,
    workbook_xml: &[u8],
    rels_xml: &[u8],
    content_types_xml: &[u8],
) -> AppResult<Vec<u8>> {
    if entries.iter().any(|e| e.name == new_name) {
        bail!("Sheet already exists: {new_name}");
    }

    let file_idx = next_sheet_file_index(&mut archive);
    let sheet_id = next_sheet_id(entries);
    let new_path = format!("xl/worksheets/sheet{file_idx}.xml");
    let new_rid = next_free_rid(rels_xml)?;

    let empty_sheet_xml = crate::xlsx_write::render_sheet_xml(&crate::save_logic::empty_sheet())?;

    let new_workbook_xml = add_sheet_to_workbook_xml(workbook_xml, new_name, &new_rid, sheet_id)?;
    let new_rels_xml = add_relationship_to_rels(rels_xml, &new_rid, &format!("worksheets/sheet{file_idx}.xml"))?;
    let new_content_types_xml =
        add_override_to_content_types(content_types_xml, &format!("/xl/worksheets/sheet{file_idx}.xml"))?;

    write_xlsx_with_replacements(
        &mut archive,
        &[
            ("xl/workbook.xml", new_workbook_xml),
            ("xl/_rels/workbook.xml.rels", new_rels_xml),
            ("[Content_Types].xml", new_content_types_xml),
        ],
        &[(new_path, empty_sheet_xml)],
    )
}

/// Đổi tên sheet: chỉ sửa attribute name= trong workbook.xml.
pub fn rename_sheet<R: Read + Seek>(
    mut archive: ZipArchive<R>,
    entries: &[SheetEntry],
    old_name: &str,
    new_name: &str,
    workbook_xml: &[u8],
) -> AppResult<Vec<u8>> {
    let target = entries
        .iter()
        .find(|e| e.name == old_name)
        .ok_or_else(|| crate::error::AppError(format!("Sheet not found: {old_name}")))?;

    if entries.iter().any(|e| e.name == new_name) {
        bail!("Sheet already exists: {new_name}");
    }

    let new_workbook_xml =
        crate::xlsx_write::rename_sheet_in_workbook_xml(workbook_xml, &target.rid, new_name)?;

    write_xlsx_with_replacements(
        &mut archive,
        &[("xl/workbook.xml", new_workbook_xml)],
        &[],
    )
}

/// Xoá sheet: gỡ <sheet> khỏi workbook.xml, gỡ <Relationship> khỏi rels, gỡ
/// <Override> khỏi Content_Types, và bỏ luôn entry sheetN.xml khỏi zip.
/// Không cho xoá nếu là sheet cuối cùng.
pub fn delete_sheet<R: Read + Seek>(
    mut archive: ZipArchive<R>,
    entries: &[SheetEntry],
    target_name: &str,
    workbook_xml: &[u8],
    rels_xml: &[u8],
    content_types_xml: &[u8],
) -> AppResult<Vec<u8>> {
    if entries.len() == 1 {
        bail!("Cannot delete last sheet");
    }
    let target = entries
        .iter()
        .find(|e| e.name == target_name)
        .ok_or_else(|| crate::error::AppError(format!("Sheet not found: {target_name}")))?;

    let new_workbook_xml = remove_sheet_from_workbook_xml(workbook_xml, &target.rid)?;
    let new_rels_xml = remove_relationship_from_rels(rels_xml, &target.rid)?;
    let target_abs_path = format!("/{}", target.path);
    let new_content_types_xml = remove_override_from_content_types(content_types_xml, &target_abs_path)?;

    write_xlsx_with_replacements_and_removals(
        &mut archive,
        &[
            ("xl/workbook.xml", new_workbook_xml),
            ("xl/_rels/workbook.xml.rels", new_rels_xml),
            ("[Content_Types].xml", new_content_types_xml),
        ],
        &[target.path.clone()],
    )
}

// ----------------------------------------------------------------------------
// Helpers XML cấp thấp
// ----------------------------------------------------------------------------

/// Sinh r:id mới chưa tồn tại trong rels.xml, dạng "rIdN".
fn next_free_rid(rels_xml: &[u8]) -> AppResult<String> {
    let mut max_n = 0u32;
    let mut reader = quick_xml::reader::Reader::from_reader(rels_xml);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Empty(e) | Event::Start(e) if e.local_name().as_ref() == b"Relationship" => {
                for a in e.attributes().filter_map(|a| a.ok()) {
                    if a.key.as_ref() == b"Id" {
                        if let Ok(val) = a.unescape_value() {
                            if let Some(num) = val.strip_prefix("rId").and_then(|s| s.parse::<u32>().ok()) {
                                max_n = max_n.max(num);
                            }
                        }
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(format!("rId{}", max_n + 1))
}

fn add_sheet_to_workbook_xml(
    xml: &[u8],
    name: &str,
    rid: &str,
    sheet_id: u32,
) -> AppResult<Vec<u8>> {
    let mut reader = quick_xml::reader::Reader::from_reader(xml);
    reader.trim_text(false);
    let mut writer = Writer::new(Vec::new());
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Eof => break,
            Event::End(e) if e.local_name().as_ref() == b"sheets" => {
                // Chèn <sheet> mới ngay trước </sheets>
                let mut new_sheet = BytesStart::new("sheet");
                new_sheet.push_attribute(("name", name));
                let sheet_id_str = sheet_id.to_string();
                new_sheet.push_attribute(("sheetId", sheet_id_str.as_str()));
                new_sheet.push_attribute(("r:id", rid));
                writer.write_event(Event::Empty(new_sheet))?;
                writer.write_event(Event::End(e))?;
            }
            other => {
                writer.write_event(other)?;
            }
        }
        buf.clear();
    }

    Ok(writer.into_inner())
}

fn remove_sheet_from_workbook_xml(xml: &[u8], target_rid: &str) -> AppResult<Vec<u8>> {
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
                if !is_target {
                    writer.write_event(Event::Empty(e))?;
                }
                // is_target == true -> bỏ qua, không ghi gì (xoá tag)
            }
            other => {
                writer.write_event(other)?;
            }
        }
        buf.clear();
    }

    Ok(writer.into_inner())
}

fn add_relationship_to_rels(xml: &[u8], rid: &str, target: &str) -> AppResult<Vec<u8>> {
    let mut reader = quick_xml::reader::Reader::from_reader(xml);
    reader.trim_text(false);
    let mut writer = Writer::new(Vec::new());
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Eof => break,
            Event::End(e) if e.local_name().as_ref() == b"Relationships" => {
                let mut rel = BytesStart::new("Relationship");
                rel.push_attribute(("Id", rid));
                rel.push_attribute((
                    "Type",
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
                ));
                rel.push_attribute(("Target", target));
                writer.write_event(Event::Empty(rel))?;
                writer.write_event(Event::End(e))?;
            }
            other => {
                writer.write_event(other)?;
            }
        }
        buf.clear();
    }

    Ok(writer.into_inner())
}

fn remove_relationship_from_rels(xml: &[u8], target_rid: &str) -> AppResult<Vec<u8>> {
    let mut reader = quick_xml::reader::Reader::from_reader(xml);
    reader.trim_text(false);
    let mut writer = Writer::new(Vec::new());
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Eof => break,
            Event::Empty(e) if e.local_name().as_ref() == b"Relationship" => {
                let is_target = e
                    .attributes()
                    .filter_map(|a| a.ok())
                    .any(|a| a.key.as_ref() == b"Id" && a.unescape_value().map(|v| v == target_rid).unwrap_or(false));
                if !is_target {
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

fn add_override_to_content_types(xml: &[u8], part_name: &str) -> AppResult<Vec<u8>> {
    let mut reader = quick_xml::reader::Reader::from_reader(xml);
    reader.trim_text(false);
    let mut writer = Writer::new(Vec::new());
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Eof => break,
            Event::End(e) if e.local_name().as_ref() == b"Types" => {
                let mut ov = BytesStart::new("Override");
                ov.push_attribute(("PartName", part_name));
                ov.push_attribute((
                    "ContentType",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
                ));
                writer.write_event(Event::Empty(ov))?;
                writer.write_event(Event::End(e))?;
            }
            other => {
                writer.write_event(other)?;
            }
        }
        buf.clear();
    }

    Ok(writer.into_inner())
}

fn remove_override_from_content_types(xml: &[u8], part_name: &str) -> AppResult<Vec<u8>> {
    let mut reader = quick_xml::reader::Reader::from_reader(xml);
    reader.trim_text(false);
    let mut writer = Writer::new(Vec::new());
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Eof => break,
            Event::Empty(e) if e.local_name().as_ref() == b"Override" => {
                let is_target = e.attributes().filter_map(|a| a.ok()).any(|a| {
                    a.key.as_ref() == b"PartName"
                        && a.unescape_value().map(|v| v == part_name).unwrap_or(false)
                });
                if !is_target {
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

/// Ghi lại zip: thay thế nội dung 1 số entry đã có (`replacements`), thêm
/// các entry hoàn toàn mới (`additions`), giữ nguyên mọi entry khác.
fn write_xlsx_with_replacements<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    replacements: &[(&str, Vec<u8>)],
    additions: &[(String, Vec<u8>)],
) -> AppResult<Vec<u8>> {
    write_xlsx_internal(archive, replacements, additions, &[])
}

fn write_xlsx_with_replacements_and_removals<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    replacements: &[(&str, Vec<u8>)],
    removals: &[String],
) -> AppResult<Vec<u8>> {
    write_xlsx_internal(archive, replacements, &[], removals)
}

fn write_xlsx_internal<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    replacements: &[(&str, Vec<u8>)],
    additions: &[(String, Vec<u8>)],
    removals: &[String],
) -> AppResult<Vec<u8>> {
    let mut out_buf = Vec::new();
    {
        let cursor = std::io::Cursor::new(&mut out_buf);
        let mut zw = ZipWriter::new(cursor);
        let options = FileOptions::default().compression_method(zip::CompressionMethod::Deflated);

        for i in 0..archive.len() {
            let mut entry = archive.by_index(i)?;
            let name = entry.name().to_string();

            if removals.iter().any(|r| r == &name) {
                continue;
            }

            zw.start_file(&name, options)?;
            if let Some((_, new_bytes)) = replacements.iter().find(|(p, _)| *p == name) {
                zw.write_all(new_bytes)?;
            } else {
                let mut buf = Vec::with_capacity(entry.size() as usize);
                entry.read_to_end(&mut buf)?;
                zw.write_all(&buf)?;
            }
        }

        for (path, bytes) in additions {
            zw.start_file(path, options)?;
            zw.write_all(bytes)?;
        }

        zw.finish()?;
    }
    Ok(out_buf)
}

/// Đọc toàn bộ workbook.xml / rels / content types cùng lúc, dùng chung bởi
/// các hàm quản lý sheet ở trên.
pub fn read_workbook_parts<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
) -> AppResult<(Vec<u8>, Vec<u8>, Vec<u8>)> {
    let workbook_xml = read_zip_entry(archive, "xl/workbook.xml")?;
    let rels_xml = read_zip_entry(archive, "xl/_rels/workbook.xml.rels")?;
    let content_types_xml = read_zip_entry(archive, "[Content_Types].xml")?;
    Ok((workbook_xml, rels_xml, content_types_xml))
}
