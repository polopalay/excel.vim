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
use model::{Rgb, SheetEntry};

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
        "setbg" => {
            let path = require_arg(args, 2, "xlsx_file")?;
            let range_ref = require_arg(args, 3, "cell_or_range")?;
            let color = require_arg(args, 4, "color")?;
            let sheet = args.get(5).map(|s| s.as_str());
            cmd_apply_style(path, range_ref, StyleAction::SetBg(parse_color(color)?), sheet)
        }
        "setfg" => {
            let path = require_arg(args, 2, "xlsx_file")?;
            let range_ref = require_arg(args, 3, "cell_or_range")?;
            let color = require_arg(args, 4, "color")?;
            let sheet = args.get(5).map(|s| s.as_str());
            cmd_apply_style(path, range_ref, StyleAction::SetFg(parse_color(color)?), sheet)
        }
        "togglebold" => {
            let path = require_arg(args, 2, "xlsx_file")?;
            let range_ref = require_arg(args, 3, "cell_or_range")?;
            let sheet = args.get(4).map(|s| s.as_str());
            cmd_apply_style(path, range_ref, StyleAction::ToggleBold, sheet)
        }
        "toggleitalic" => {
            let path = require_arg(args, 2, "xlsx_file")?;
            let range_ref = require_arg(args, 3, "cell_or_range")?;
            let sheet = args.get(4).map(|s| s.as_str());
            cmd_apply_style(path, range_ref, StyleAction::ToggleItalic, sheet)
        }
        "merge" => {
            let path = require_arg(args, 2, "xlsx_file")?;
            let range_ref = require_arg(args, 3, "range")?;
            let sheet = args.get(4).map(|s| s.as_str());
            cmd_merge(path, range_ref, sheet)
        }
        "unmerge" => {
            let path = require_arg(args, 2, "xlsx_file")?;
            let range_ref = require_arg(args, 3, "cell_or_range")?;
            let sheet = args.get(4).map(|s| s.as_str());
            cmd_unmerge(path, range_ref, sheet)
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

/// Đọc xl/styles.xml -> Vec<CellStyle> đã resolve theo index cellXfs.
/// File có thể không có styles.xml (rất hiếm) -> coi như chỉ có style mặc định.
fn load_styles<R: std::io::Read + std::io::Seek>(
    archive: &mut ZipArchive<R>,
) -> AppResult<Vec<model::CellStyle>> {
    let bytes = match xlsx_read::read_zip_entry(archive, "xl/styles.xml") {
        Ok(b) => Some(b),
        Err(_) => None,
    };
    xlsx_read::read_styles(bytes.as_deref())
}

/// Tên màu chuẩn (tiếng Anh, khớp gợi ý completion trong Vim) gợi nhớ, đủ
/// dùng cho các lệnh style mà không cần nhớ mã hex. Vẫn hỗ trợ alias tiếng
/// Việt cũ (đỏ, vàng, xanh-la...) để tương thích ngược. Khoảng trắng/gạch
/// nối trong tên màu đều được bỏ qua khi so khớp.
fn parse_color(input: &str) -> AppResult<Option<Rgb>> {
    let lower = input.trim().to_lowercase();
    let normalized: String = lower.chars().filter(|c| *c != ' ' && *c != '-' && *c != '_').collect();
    if normalized == "none" || normalized == "clear" || normalized == "xoa" || normalized == "xóa" {
        return Ok(None);
    }
    let named = match normalized.as_str() {
        "red" | "do" | "đỏ" => Some(Rgb { r: 0xFF, g: 0x00, b: 0x00 }),
        "green" | "xanhla" | "la" | "lá" => Some(Rgb { r: 0x00, g: 0xB0, b: 0x50 }),
        "blue" | "xanhduong" | "duong" | "dương" => Some(Rgb { r: 0x00, g: 0x70, b: 0xC0 }),
        "yellow" | "vang" | "vàng" => Some(Rgb { r: 0xFF, g: 0xFF, b: 0x00 }),
        "orange" | "cam" => Some(Rgb { r: 0xFF, g: 0xA5, b: 0x00 }),
        "purple" | "tim" | "tím" => Some(Rgb { r: 0x80, g: 0x00, b: 0x80 }),
        "gray" | "grey" | "xam" | "xám" => Some(Rgb { r: 0xD9, g: 0xD9, b: 0xD9 }),
        "white" | "trang" | "trắng" => Some(Rgb { r: 0xFF, g: 0xFF, b: 0xFF }),
        "black" | "den" | "đen" => Some(Rgb { r: 0x00, g: 0x00, b: 0x00 }),
        _ => None,
    };
    if let Some(c) = named {
        return Ok(Some(c));
    }
    Rgb::from_hex(&lower)
        .map(Some)
        .ok_or_else(|| AppError(format!("Màu không hợp lệ: {input} (dùng tên màu hoặc #RRGGBB)")))
}

/// Parse 1 trong các dạng địa chỉ cell sau, trả về toàn bộ toạ độ (col, row):
/// - "B2"           : 1 cell đơn
/// - "B2:D5"        : 1 vùng hình chữ nhật
/// - "B2,C2,B3,C3"  : danh sách rời rạc (cách nhau bởi dấu phẩy), dùng khi
///                    Vim đã resolve sẵn list cell trong Visual selection.
///
/// Trùng lặp được khử (cùng 1 (col,row) không xuất hiện 2 lần) để tránh áp
/// style 2 lần trên cùng cell — toggle bold/italic sẽ bị huỷ nhau nếu lặp.
fn parse_cell_or_range(input: &str) -> AppResult<Vec<(u32, u32)>> {
    use std::collections::BTreeSet;
    let mut seen: BTreeSet<(u32, u32)> = BTreeSet::new();
    let mut coords = Vec::new();

    for part in input.split(',') {
        let part = part.trim();
        if part.is_empty() {
            continue;
        }
        if let Some((start, end)) = part.split_once(':') {
            let (c1, r1, c2, r2) = coord::parse_range(&format!("{start}:{end}"))?;
            for r in r1..=r2 {
                for c in c1..=c2 {
                    if seen.insert((c, r)) {
                        coords.push((c, r));
                    }
                }
            }
        } else {
            let (col, row) = coord::ref_to_col_row(part)?;
            if seen.insert((col, row)) {
                coords.push((col, row));
            }
        }
    }
    if coords.is_empty() {
        return Err(AppError(format!("Không có cell hợp lệ: {input}")));
    }
    Ok(coords)
}

/// Hành động style mà 4 lệnh Vim (:ExcelSetBg/:ExcelSetFg/:ExcelBold/
/// :ExcelItalic) có thể yêu cầu, áp dụng đồng nhất cho 1 cell hoặc 1 vùng.
enum StyleAction {
    SetBg(Option<Rgb>),
    SetFg(Option<Rgb>),
    ToggleBold,
    ToggleItalic,
}

impl StyleAction {
    fn apply(&self, sheet: &mut model::SheetData, row: u32, col: u32) {
        match self {
            StyleAction::SetBg(rgb) => sheet.set_bg_color(row, col, *rgb),
            StyleAction::SetFg(rgb) => sheet.set_fg_color(row, col, *rgb),
            StyleAction::ToggleBold => sheet.toggle_bold(row, col),
            StyleAction::ToggleItalic => sheet.toggle_italic(row, col),
        }
    }
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
    let styles = load_styles(&mut archive)?;

    let sheet_xml = xlsx_read::read_zip_entry(&mut archive, &target.path)?;
    let sheet = xlsx_read::parse_sheet_xml(&sheet_xml, &shared, &styles)?;

    print_table_with_style_metadata(&sheet);
    Ok(())
}

/// In bảng ASCII ra stdout, theo sau bởi 2 khối metadata:
///   1. @@STYLE@@: style hiển thị (bold/italic/màu) cho các cell có style
///      khác mặc định — dùng để highlight.
///   2. @@CELLMAP@@: map TOÀN BỘ vị trí buffer -> địa chỉ cell Excel gốc —
///      dùng để file .vim xác định cursor/Visual selection đang ở cell nào
///      (cần cho :ExcelSetBg/:ExcelSetFg/:ExcelBold/:ExcelItalic khi gọi
///      không kèm tham số cell, lấy từ vị trí con trỏ).
///
/// Format khối @@STYLE@@ (mỗi dòng 1 cell có style khác mặc định):
///   <line>\t<col_start>\t<col_end>\t<bold>\t<italic>\t<font_hex|->\t<bg_hex|->
/// Format khối @@CELLMAP@@ (mỗi dòng 1 cell hiển thị, kể cả cell rỗng):
///   <line>\t<col_start>\t<col_end>\t<cell_ref>
/// - line: số dòng buffer (1-based, tính cả dòng viền +---+)
/// - col_start/col_end: cột ký tự (1-based, inclusive) trong dòng buffer.
/// - cell_ref: địa chỉ Excel của TOP-LEFT vùng merge (hoặc chính cell đó
///   nếu không merge), ví dụ "B2" — vì style/chỉnh sửa luôn áp dụng ở mức
///   top-left cho vùng merge.
fn print_table_with_style_metadata(sheet: &model::SheetData) {
    let (display_rows, coords_rows, vmerge_below, col_spans) = display::build_display_rows_full(sheet);
    let (text, spans, full_spans) = table::render_with_spans(&display_rows, &coords_rows, &col_spans, &vmerge_below);
    println!("{text}");

    println!("@@STYLE@@");
    for (row_idx, (coord_row, span_row)) in coords_rows.iter().zip(spans.iter()).enumerate() {
        let buffer_line = row_idx * 2 + 2;
        for ((row, col), (start, end)) in coord_row.iter().zip(span_row.iter()) {
            let (tl_row, tl_col) = display::merge_top_left(&sheet.merges, *row, *col);
            let style = sheet.get_style(tl_row, tl_col);
            if !style.bold && !style.italic && style.font_color.is_none() && style.bg_color.is_none() {
                continue;
            }
            let font_hex = style.font_color.map(|c| c.to_hex()).unwrap_or_else(|| "-".to_string());
            let bg_hex = style.bg_color.map(|c| c.to_hex()).unwrap_or_else(|| "-".to_string());
            // Cell rỗng (end<=start): không có ký tự nội dung để highlight,
            // nhưng vẫn cần hiển thị màu nền — highlight đúng 1 ký tự
            // padding (space) tại vị trí start để background vẫn thấy được.
            let end_incl = if *end > *start { *end - 1 } else { *start };
            println!(
                "{buffer_line}\t{start}\t{end_incl}\t{bold}\t{italic}\t{font_hex}\t{bg_hex}",
                bold = style.bold as u8,
                italic = style.italic as u8,
            );
        }
    }
    println!("@@END@@");

    println!("@@CELLMAP@@");
    // Dùng full_spans (vị trí của TOÀN cell, bao gồm padding) để Vim biết
    // cell merge ngang "phủ" tới đâu trên dòng — quan trọng cho Visual block
    // select dọc: khi user Ctrl-V chọn 1 cột, các cell merge ngang chứa cột
    // đó (vd A1 trong merge A1:D1) phải được include vào selection.
    for (row_idx, (coord_row, full_span_row)) in coords_rows.iter().zip(full_spans.iter()).enumerate() {
        let buffer_line = row_idx * 2 + 2;
        for ((row, col), (full_start, full_end)) in coord_row.iter().zip(full_span_row.iter()) {
            let (tl_row, tl_col) = display::merge_top_left(&sheet.merges, *row, *col);
            let cell_ref = coord::col_row_to_ref(tl_col, tl_row);
            let end_incl = if *full_end > *full_start { *full_end - 1 } else { *full_start };
            println!("{buffer_line}\t{full_start}\t{end_incl}\t{cell_ref}");
        }
    }
    println!("@@CELLMAPEND@@");
}

fn cmd_save(path: &str, txt_file: &str, sheet_name: Option<&str>) -> AppResult<()> {
    ensure_workbook(path)?;
    let mut archive = open_archive(path)?;
    let entries = load_sheet_entries(&mut archive)?;
    let target = find_sheet(&entries, sheet_name)?.clone();
    let shared = load_shared_strings(&mut archive)?;
    let styles = load_styles(&mut archive)?;

    let sheet_xml = xlsx_read::read_zip_entry(&mut archive, &target.path)?;
    let original_sheet = xlsx_read::parse_sheet_xml(&sheet_xml, &shared, &styles)?;

    let txt_content = fs::read_to_string(txt_file)
        .map_err(|e| AppError(format!("Cannot read txt_file {txt_file}: {e}")))?;
    let lines: Vec<String> = txt_content.lines().map(|s| s.to_string()).collect();

    let new_sheet = save_logic::apply_ascii_table(&original_sheet, &lines)?;
    let new_sheet_xml = xlsx_write::render_sheet_xml(&new_sheet)?;

    // Nếu user đã thêm style mới (qua setbg) chưa tồn tại trong styles.xml
    // gốc, đồng bộ lại styles.xml để index s="N" mới ghi trong sheet XML
    // có entry tương ứng. Style cũ giữ nguyên bytes-for-bytes.
    let original_styles_xml = xlsx_read::read_zip_entry(&mut archive, "xl/styles.xml").ok();
    let synced_styles_xml = xlsx_write::sync_styles_xml(original_styles_xml.as_deref(), &new_sheet.styles)?;

    let mut replacements: Vec<(&str, Vec<u8>)> = vec![(target.path.as_str(), new_sheet_xml)];
    if let Some(new_styles_xml) = synced_styles_xml {
        replacements.push(("xl/styles.xml", new_styles_xml));
    }

    let new_xlsx_bytes = xlsx_write::write_xlsx_replacing_entries(&mut archive, &replacements)?;

    atomic_write(path, &new_xlsx_bytes)?;
    Ok(())
}

/// Áp dụng 1 style action (đổi màu nền/màu chữ/đảo bold/đảo italic) cho 1
/// cell hoặc 1 vùng (range, ví dụ "B2:D5") trong sheet, rồi ghi lại file.
/// Dùng chung cho cả 4 lệnh Vim :ExcelSetBg/:ExcelSetFg/:ExcelBold/:ExcelItalic.
fn cmd_apply_style(
    path: &str,
    range_ref: &str,
    action: StyleAction,
    sheet_name: Option<&str>,
) -> AppResult<()> {
    ensure_workbook(path)?;
    let mut archive = open_archive(path)?;
    let entries = load_sheet_entries(&mut archive)?;
    let target = find_sheet(&entries, sheet_name)?.clone();
    let shared = load_shared_strings(&mut archive)?;
    let styles = load_styles(&mut archive)?;

    let sheet_xml = xlsx_read::read_zip_entry(&mut archive, &target.path)?;
    let mut sheet = xlsx_read::parse_sheet_xml(&sheet_xml, &shared, &styles)?;

    let coords = parse_cell_or_range(range_ref)?;
    for (col, row) in coords {
        action.apply(&mut sheet, row, col);
    }

    let new_sheet_xml = xlsx_write::render_sheet_xml(&sheet)?;
    let original_styles_xml = xlsx_read::read_zip_entry(&mut archive, "xl/styles.xml").ok();
    let synced_styles_xml = xlsx_write::sync_styles_xml(original_styles_xml.as_deref(), &sheet.styles)?;

    let mut replacements: Vec<(&str, Vec<u8>)> = vec![(target.path.as_str(), new_sheet_xml)];
    if let Some(new_styles_xml) = synced_styles_xml {
        replacements.push(("xl/styles.xml", new_styles_xml));
    }

    let new_xlsx_bytes = xlsx_write::write_xlsx_replacing_entries(&mut archive, &replacements)?;
    atomic_write(path, &new_xlsx_bytes)?;
    Ok(())
}

/// Tính bounding box của 1 chuỗi input range/list ("B2:D5" hoặc "B2,C3,D5"
/// hoặc mix) — tức tìm (min_row, min_col, max_row, max_col) bao trùm mọi
/// cell được liệt kê. Dùng cho lệnh merge (vùng gộp là hình chữ nhật) và
/// unmerge (vùng tìm kiếm các merge giao với input).
fn parse_bounding_box(input: &str) -> AppResult<(u32, u32, u32, u32)> {
    let coords = parse_cell_or_range(input)?;
    let (mut min_r, mut min_c, mut max_r, mut max_c) = (u32::MAX, u32::MAX, 0u32, 0u32);
    for (col, row) in coords {
        if row < min_r {
            min_r = row;
        }
        if row > max_r {
            max_r = row;
        }
        if col < min_c {
            min_c = col;
        }
        if col > max_c {
            max_c = col;
        }
    }
    Ok((min_r, min_c, max_r, max_c))
}

/// Gộp 1 vùng các ô (giống nút Merge Cells trong Excel). Vùng được truyền
/// dưới dạng "B2:D5" hoặc danh sách comma-separated/range mix — sẽ tính
/// bounding box bao trùm.
fn cmd_merge(path: &str, range_ref: &str, sheet_name: Option<&str>) -> AppResult<()> {
    ensure_workbook(path)?;
    let mut archive = open_archive(path)?;
    let entries = load_sheet_entries(&mut archive)?;
    let target = find_sheet(&entries, sheet_name)?.clone();
    let shared = load_shared_strings(&mut archive)?;
    let styles = load_styles(&mut archive)?;

    let sheet_xml = xlsx_read::read_zip_entry(&mut archive, &target.path)?;
    let mut sheet = xlsx_read::parse_sheet_xml(&sheet_xml, &shared, &styles)?;

    let (min_r, min_c, max_r, max_c) = parse_bounding_box(range_ref)?;
    if !sheet.add_merge(min_r, min_c, max_r, max_c) {
        return Err(AppError(format!(
            "Vùng không hợp lệ để gộp: {range_ref} (cần ít nhất 2 ô)"
        )));
    }

    let new_sheet_xml = xlsx_write::render_sheet_xml(&sheet)?;
    let original_styles_xml = xlsx_read::read_zip_entry(&mut archive, "xl/styles.xml").ok();
    let synced_styles_xml = xlsx_write::sync_styles_xml(original_styles_xml.as_deref(), &sheet.styles)?;

    let mut replacements: Vec<(&str, Vec<u8>)> = vec![(target.path.as_str(), new_sheet_xml)];
    if let Some(new_styles_xml) = synced_styles_xml {
        replacements.push(("xl/styles.xml", new_styles_xml));
    }

    let new_xlsx_bytes = xlsx_write::write_xlsx_replacing_entries(&mut archive, &replacements)?;
    atomic_write(path, &new_xlsx_bytes)?;
    Ok(())
}

/// Bỏ gộp mọi vùng merge giao với 1 cell hoặc 1 range cho trước. Nếu input
/// là 1 cell đơn (ví dụ "B2"), tìm merge chứa cell đó. Nếu input là range
/// ("B2:D5") hoặc list, bỏ mọi merge giao với bounding box.
fn cmd_unmerge(path: &str, range_ref: &str, sheet_name: Option<&str>) -> AppResult<()> {
    ensure_workbook(path)?;
    let mut archive = open_archive(path)?;
    let entries = load_sheet_entries(&mut archive)?;
    let target = find_sheet(&entries, sheet_name)?.clone();
    let shared = load_shared_strings(&mut archive)?;
    let styles = load_styles(&mut archive)?;

    let sheet_xml = xlsx_read::read_zip_entry(&mut archive, &target.path)?;
    let mut sheet = xlsx_read::parse_sheet_xml(&sheet_xml, &shared, &styles)?;

    let (min_r, min_c, max_r, max_c) = parse_bounding_box(range_ref)?;
    let removed = sheet.remove_merges_intersecting(min_r, min_c, max_r, max_c);
    if removed == 0 {
        return Err(AppError(format!("Không có vùng gộp nào trong {range_ref}")));
    }

    let new_sheet_xml = xlsx_write::render_sheet_xml(&sheet)?;
    let original_styles_xml = xlsx_read::read_zip_entry(&mut archive, "xl/styles.xml").ok();
    let synced_styles_xml = xlsx_write::sync_styles_xml(original_styles_xml.as_deref(), &sheet.styles)?;

    let mut replacements: Vec<(&str, Vec<u8>)> = vec![(target.path.as_str(), new_sheet_xml)];
    if let Some(new_styles_xml) = synced_styles_xml {
        replacements.push(("xl/styles.xml", new_styles_xml));
    }

    let new_xlsx_bytes = xlsx_write::write_xlsx_replacing_entries(&mut archive, &replacements)?;
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
