mod coord;
mod display;
mod error;
mod model;
mod save_logic;
mod sheet_ops;
mod table;
mod xlsx_read;
mod xlsx_write;

use std::fs;
use std::io::Cursor;
use std::path::Path;

use zip::ZipArchive;

use error::{AppError, AppResult};
use model::SheetEntry;

fn main() {
    let args: Vec<String> = std::env::args().collect();
    if let Err(e) = run(&args) {
        eprintln!("{e}");
        std::process::exit(1);
    }
}

fn run(args: &[String]) -> AppResult<()> {
    let mode = args
        .get(1)
        .ok_or_else(|| AppError("Missing mode argument".to_string()))?;

    match mode.as_str() {
        "open" => {
            let path = require_arg(args, 2, "xlsx_file")?;
            let sheet = args.get(3).map(|s| s.as_str());
            cmd_open(path, sheet)
        }
        "save" => {
            let path = require_arg(args, 2, "xlsx_file")?;
            let txt_file = require_arg(args, 3, "txt_file")?;
            let sheet = args.get(4).map(|s| s.as_str());
            cmd_save(path, txt_file, sheet)
        }
        "sheets" => {
            let path = require_arg(args, 2, "xlsx_file")?;
            cmd_sheets(path)
        }
        "addsheet" => {
            let path = require_arg(args, 2, "xlsx_file")?;
            let name = require_arg(args, 3, "sheet_name")?;
            cmd_addsheet(path, name)
        }
        "rensheet" => {
            let path = require_arg(args, 2, "xlsx_file")?;
            let old_name = require_arg(args, 3, "old_name")?;
            let new_name = require_arg(args, 4, "new_name")?;
            cmd_rensheet(path, old_name, new_name)
        }
        "delsheet" => {
            let path = require_arg(args, 2, "xlsx_file")?;
            let name = require_arg(args, 3, "sheet_name")?;
            cmd_delsheet(path, name)
        }
        other => Err(AppError(format!("Unknown mode: {other}"))),
    }
}

fn require_arg<'a>(args: &'a [String], idx: usize, label: &str) -> AppResult<&'a str> {
    args.get(idx)
        .map(|s| s.as_str())
        .ok_or_else(|| AppError(format!("Missing argument: {label}")))
}

// ----------------------------------------------------------------------------
// ensure_workbook: tương đương Python — tạo file mới rỗng nếu chưa tồn tại
// ----------------------------------------------------------------------------

/// Nếu file chưa tồn tại, tạo 1 workbook .xlsx tối giản hợp lệ (1 sheet
/// "Sheet1", 1 ô A1 rỗng) rồi ghi ra đĩa, giống `ensure_workbook()` Python.
fn ensure_workbook(path: &str) -> AppResult<()> {
    if Path::new(path).exists() {
        return Ok(());
    }
    let bytes = build_minimal_xlsx("Sheet1")?;
    atomic_write(path, &bytes)
}

/// Sinh bytes 1 file .xlsx tối giản từ đầu (không cần file mẫu), chứa đủ
/// các phần XML bắt buộc theo chuẩn OOXML để Excel/LibreOffice mở được.
fn build_minimal_xlsx(sheet_name: &str) -> AppResult<Vec<u8>> {
    use std::io::Write;
    use zip::write::FileOptions;
    use zip::ZipWriter;

    let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>"#;

    let root_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#;

    let workbook_xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="{sheet_name}" sheetId="1" r:id="rId1"/></sheets>
</workbook>"#
    );

    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#;

    let sheet_xml = xlsx_write::render_sheet_xml(&save_logic::empty_sheet())?;

    let mut buf = Vec::new();
    {
        let cursor = Cursor::new(&mut buf);
        let mut zw = ZipWriter::new(cursor);
        let options = FileOptions::default().compression_method(zip::CompressionMethod::Deflated);

        zw.start_file("[Content_Types].xml", options)?;
        zw.write_all(content_types.as_bytes())?;

        zw.start_file("_rels/.rels", options)?;
        zw.write_all(root_rels.as_bytes())?;

        zw.start_file("xl/workbook.xml", options)?;
        zw.write_all(workbook_xml.as_bytes())?;

        zw.start_file("xl/_rels/workbook.xml.rels", options)?;
        zw.write_all(workbook_rels.as_bytes())?;

        zw.start_file("xl/worksheets/sheet1.xml", options)?;
        zw.write_all(&sheet_xml)?;

        zw.finish()?;
    }
    Ok(buf)
}

/// Ghi file atomic: viết ra file tạm cùng thư mục rồi rename, tránh để file
/// .xlsx ở trạng thái hỏng nếu chương trình bị ngắt giữa lúc ghi.
fn atomic_write(path: &str, bytes: &[u8]) -> AppResult<()> {
    let target = Path::new(path);
    let dir = target.parent().unwrap_or_else(|| Path::new("."));
    let tmp_path = dir.join(format!(
        ".{}.tmp{}",
        target.file_name().and_then(|n| n.to_str()).unwrap_or("excel_rs"),
        std::process::id()
    ));
    fs::write(&tmp_path, bytes)?;
    fs::rename(&tmp_path, target)?;
    Ok(())
}

fn open_archive(path: &str) -> AppResult<ZipArchive<std::io::BufReader<fs::File>>> {
    let file = fs::File::open(path)?;
    let reader = std::io::BufReader::new(file);
    Ok(ZipArchive::new(reader)?)
}

fn load_sheet_entries<R: std::io::Read + std::io::Seek>(
    archive: &mut ZipArchive<R>,
) -> AppResult<Vec<SheetEntry>> {
    let workbook_xml = xlsx_read::read_zip_entry(archive, "xl/workbook.xml")?;
    let rels_xml = xlsx_read::read_zip_entry(archive, "xl/_rels/workbook.xml.rels")?;
    xlsx_read::read_sheet_entries(&workbook_xml, &rels_xml)
}

fn find_sheet<'a>(entries: &'a [SheetEntry], name: Option<&str>) -> AppResult<&'a SheetEntry> {
    match name {
        Some(n) if !n.is_empty() => entries
            .iter()
            .find(|e| e.name == n)
            .ok_or_else(|| AppError(format!("Sheet not found: {n}"))),
        _ => entries
            .first()
            .ok_or_else(|| AppError("No sheets found".to_string())),
    }
}

fn load_shared_strings<R: std::io::Read + std::io::Seek>(
    archive: &mut ZipArchive<R>,
) -> AppResult<Vec<String>> {
    let bytes = match xlsx_read::read_zip_entry(archive, "xl/sharedStrings.xml") {
        Ok(b) => Some(b),
        Err(_) => None, // file có thể không có sharedStrings.xml (toàn inline string)
    };
    xlsx_read::read_shared_strings(bytes.as_deref())
}

// ----------------------------------------------------------------------------
// Commands
// ----------------------------------------------------------------------------

fn cmd_open(path: &str, sheet_name: Option<&str>) -> AppResult<()> {
    ensure_workbook(path)?;
    let mut archive = open_archive(path)?;
    let entries = load_sheet_entries(&mut archive)?;
    let target = find_sheet(&entries, sheet_name)?.clone();
    let shared = load_shared_strings(&mut archive)?;

    let sheet_xml = xlsx_read::read_zip_entry(&mut archive, &target.path)?;
    let sheet = xlsx_read::parse_sheet_xml(&sheet_xml, &shared)?;

    let display_rows = display::build_display_rows(&sheet);
    println!("{}", table::render(&display_rows));
    Ok(())
}

fn cmd_save(path: &str, txt_file: &str, sheet_name: Option<&str>) -> AppResult<()> {
    ensure_workbook(path)?;
    let mut archive = open_archive(path)?;
    let entries = load_sheet_entries(&mut archive)?;
    let target = find_sheet(&entries, sheet_name)?.clone();
    let shared = load_shared_strings(&mut archive)?;

    let sheet_xml = xlsx_read::read_zip_entry(&mut archive, &target.path)?;
    let original_sheet = xlsx_read::parse_sheet_xml(&sheet_xml, &shared)?;

    let txt_content = fs::read_to_string(txt_file)
        .map_err(|e| AppError(format!("Cannot read txt_file {txt_file}: {e}")))?;
    let lines: Vec<String> = txt_content.lines().map(|s| s.to_string()).collect();

    let new_sheet = save_logic::apply_ascii_table(&original_sheet, &lines)?;
    let new_sheet_xml = xlsx_write::render_sheet_xml(&new_sheet)?;

    let new_xlsx_bytes =
        xlsx_write::write_xlsx_replacing_sheet(&mut archive, &target.path, &new_sheet_xml)?;

    atomic_write(path, &new_xlsx_bytes)?;
    Ok(())
}

fn cmd_sheets(path: &str) -> AppResult<()> {
    ensure_workbook(path)?;
    let mut archive = open_archive(path)?;
    let entries = load_sheet_entries(&mut archive)?;
    for name in sheet_ops::list_sheet_names(&entries) {
        println!("{name}");
    }
    Ok(())
}

fn cmd_addsheet(path: &str, name: &str) -> AppResult<()> {
    ensure_workbook(path)?;
    let mut archive = open_archive(path)?;
    let entries = load_sheet_entries(&mut archive)?;
    let (workbook_xml, rels_xml, content_types_xml) =
        sheet_ops::read_workbook_parts(&mut archive)?;

    let new_bytes = sheet_ops::create_sheet(
        archive,
        &entries,
        name,
        &workbook_xml,
        &rels_xml,
        &content_types_xml,
    )?;
    atomic_write(path, &new_bytes)?;
    Ok(())
}

fn cmd_rensheet(path: &str, old_name: &str, new_name: &str) -> AppResult<()> {
    ensure_workbook(path)?;
    let mut archive = open_archive(path)?;
    let entries = load_sheet_entries(&mut archive)?;
    let workbook_xml = xlsx_read::read_zip_entry(&mut archive, "xl/workbook.xml")?;

    let new_bytes = sheet_ops::rename_sheet(archive, &entries, old_name, new_name, &workbook_xml)?;
    atomic_write(path, &new_bytes)?;
    Ok(())
}

fn cmd_delsheet(path: &str, name: &str) -> AppResult<()> {
    ensure_workbook(path)?;
    let mut archive = open_archive(path)?;
    let entries = load_sheet_entries(&mut archive)?;
    let (workbook_xml, rels_xml, content_types_xml) =
        sheet_ops::read_workbook_parts(&mut archive)?;

    let new_bytes = sheet_ops::delete_sheet(
        archive,
        &entries,
        name,
        &workbook_xml,
        &rels_xml,
        &content_types_xml,
    )?;
    atomic_write(path, &new_bytes)?;
    Ok(())
}
