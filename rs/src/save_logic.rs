//! `save_xlsx` — port 1:1 logic từ excel.py, dùng SheetData ở mức XML thay
//! vì openpyxl object model.
//!
//! Một khác biệt có chủ ý so với Python: Python dùng openpyxl để
//! insert_rows/insert_cols/delete_rows/delete_cols rồi merge lại; ở đây vì
//! ta tự quản lý SheetData (map (row,col) -> value), "đồng bộ số dòng/cột"
//! chỉ cần set `max_row`/`max_col` mới — không cần thật sự dịch chuyển dữ
//! liệu, vì dữ liệu được VIẾT LẠI HOÀN TOÀN từ `new_rows` (không giữ lại ô
//! cũ ngoài phạm vi mới). Điều này khớp đúng ý nghĩa cuối cùng của code
//! Python: nội dung file sau khi save = đúng những gì user thấy trong buffer
//! Vim, với merge được khôi phục lại từ thông tin gốc.

use crate::display::{build_display_plan, col_map_per_row};
use crate::error::AppResult;
use crate::model::SheetData;
use crate::table::parse_ascii;

/// Áp dụng nội dung bảng ASCII (đã user sửa trong Vim) lên SheetData GỐC,
/// dùng đúng `col_map_per_row` tính từ cấu trúc merge HIỆN TẠI của sheet đó
/// — tương đương bước 2-4 của `save_xlsx()` trong Python.
///
/// Trả về SheetData mới (giá trị mới + merges gốc được áp lại) để ghi ra XML.
pub fn apply_ascii_table(
    original: &SheetData,
    ascii_lines: &[String],
) -> AppResult<SheetData> {
    let new_rows = parse_ascii(ascii_lines);

    // Thông tin merge/col_map dựa trên cấu trúc HIỆN TẠI của sheet gốc,
    // đúng như Python lấy info TRƯỚC KHI ghi đè.
    let plan = build_display_plan(&original.merges);
    let mapping = col_map_per_row(&plan, original.max_row.max(1), original.max_col.max(1));

    // Nếu số dòng mapping (tính từ file gốc) đủ khớp số dòng mới thì dùng,
    // nếu không (user đã thêm/xoá dòng thủ công ngoài insert_row) thì
    // fallback ánh xạ 1:1 — giống đúng logic Python.
    let col_map_per_row_final: Vec<Vec<u32>> = if mapping.len() >= new_rows.len() {
        mapping
    } else {
        new_rows
            .iter()
            .map(|r| (1..=r.len() as u32).collect())
            .collect()
    };

    let mut new_sheet = SheetData::default();

    for (r_idx, row) in new_rows.iter().enumerate() {
        let col_map = col_map_per_row_final
            .get(r_idx)
            .cloned()
            .unwrap_or_else(|| (1..=row.len() as u32).collect());

        for (i, value) in row.iter().enumerate() {
            let real_col = col_map.get(i).copied().unwrap_or(i as u32 + 1);
            // \\n (escape hiển thị trong bảng ASCII) -> newline thật,
            // đúng như Python: value.replace("\\n", "\n")
            let real_value = value.replace("\\n", "\n");
            let row_num = r_idx as u32 + 1;

            if let Some(formula_text) = real_value.strip_prefix('=') {
                // User gõ công thức mới (hoặc sửa lại công thức cũ) trực
                // tiếp trong bảng. Giá trị HIỂN THỊ sẽ được
                // formula::evaluate_all tính lại sau khi new_sheet dựng
                // xong — tạm set raw text vào đây để không bị coi là rỗng
                // nếu evaluate_all gặp lỗi bất ngờ.
                new_sheet
                    .formulas
                    .insert((row_num, real_col), formula_text.to_string());
                new_sheet.set(row_num, real_col, real_value.clone());
            } else if original.formulas.contains_key(&(row_num, real_col))
                && original.get(row_num, real_col).unwrap_or("") == real_value.as_str()
            {
                // QUAN TRỌNG: cell này có formula ở file gốc và user KHÔNG
                // sửa gì (bảng ASCII chỉ hiển thị giá trị cache, không hiển
                // thị "=..."). Nếu không xử lý riêng case này, mọi lần lưu
                // sẽ làm "='" -> chỉ literal -> XOÁ MẤT formula của mọi cell
                // user không động tới. Giữ lại formula cũ ở đây.
                new_sheet
                    .formulas
                    .insert((row_num, real_col), original.formulas[&(row_num, real_col)].clone());
                new_sheet.set(row_num, real_col, real_value);
            } else {
                // Literal value bình thường — kể cả trường hợp cell này
                // TRƯỚC ĐÓ có formula nhưng user đã gõ đè 1 giá trị khác
                // (không phải "="): coi như họ chủ động xoá formula, thay
                // bằng giá trị tĩnh (khớp hành vi "what you see is what
                // you get" của bảng ASCII).
                new_sheet.set(row_num, real_col, real_value);
            }
        }
    }

    // Áp lại các vùng merge gốc (không đổi), đúng như Python merge lại
    // ranges đã lấy ở bước 2.
    new_sheet.merges = original.merges.clone();

    // Giữ nguyên style (bold/italic/màu chữ/màu nền) của từng cell theo
    // đúng toạ độ (row, col) — vì save_xlsx chỉ ghi đè giá trị text, không
    // có UI nào cho phép sửa style qua bảng ASCII (style chỉ đổi qua lệnh
    // setbg riêng), nên style phải giữ y nguyên vị trí cũ.
    new_sheet.cell_style_id = original.cell_style_id.clone();
    new_sheet.styles = original.styles.clone();

    // Đồng bộ bounds: lấy max giữa dữ liệu mới và merges, nhưng KHÔNG giữ
    // max_row/max_col cũ nếu bảng mới nhỏ hơn — khớp hành vi
    // ws.delete_rows/delete_cols của Python khi new_count < old_count.
    new_sheet.recompute_bounds();
    // recompute_bounds() đã tự max() với merges, nhưng nếu new_rows ngắn
    // hơn merges cũ thì merges cũ có thể tham chiếu ra ngoài dữ liệu mới.
    // Giữ nguyên hành vi Python: merge vẫn được áp dù vùng dữ liệu nhỏ hơn
    // (try/except bỏ qua nếu lỗi) -> ở đây ta không lỗi vì model chỉ là
    // map thưa, nhưng đảm bảo max_row/max_col tối thiểu bằng số dòng mới.
    let min_row = new_rows.len() as u32;
    let min_col = new_rows.iter().map(|r| r.len() as u32).max().unwrap_or(0);
    if new_sheet.max_row < min_row {
        new_sheet.max_row = min_row;
    }
    if new_sheet.max_col < min_col {
        new_sheet.max_col = min_col;
    }

    // Tính lại mọi công thức (cũ giữ nguyên từ vòng lặp set() phía trên +
    // công thức mới user vừa nhập) SAU KHI toàn bộ cell/merge/style đã ổn
    // định, để công thức có thể tham chiếu tới bất kỳ cell nào trong sheet
    // bất kể thứ tự nhập liệu.
    crate::formula::evaluate_all(&mut new_sheet);

    Ok(new_sheet)
}

/// Dùng khi tạo sheet/workbook mới hoàn toàn (ensure_workbook trong Python):
/// 1 sheet, 1 ô A1 rỗng.
pub fn empty_sheet() -> SheetData {
    let mut s = SheetData::default();
    s.max_row = 1;
    s.max_col = 1;
    s
}
