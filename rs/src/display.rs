//! Logic xây "kế hoạch hiển thị" cho 1 sheet dựa trên các vùng merge,
//! port 1:1 từ `build_display_plan()` trong excel.py.

use std::collections::{HashMap, HashSet};

use crate::model::{MergeRange, SheetData};

pub struct DisplayPlan {
    /// (row, col) thuộc vùng merge nhưng KHÔNG phải top-left -> hiển thị rỗng.
    pub skip: HashSet<(u32, u32)>,
    /// row -> set các cột cần ẩn HOÀN TOÀN trên dòng đó (do bị merge ngang nuốt).
    pub hskip_cols_by_row: HashMap<u32, HashSet<u32>>,
}

/// Với 1 toạ độ (row, col), nếu nó thuộc 1 vùng merge thì trả về toạ độ
/// top-left của vùng đó; nếu không thuộc merge nào, trả về chính nó.
/// Dùng để lấy đúng style hiển thị cho mọi ô trong vùng merge (Excel chỉ
/// lưu style ở ô top-left, các ô bị merge khác thường không có style riêng
/// hoặc style không quan trọng vì nội dung luôn rỗng/ẩn).
pub fn merge_top_left(ranges: &[MergeRange], row: u32, col: u32) -> (u32, u32) {
    for rg in ranges {
        if row >= rg.min_row && row <= rg.max_row && col >= rg.min_col && col <= rg.max_col {
            return (rg.min_row, rg.min_col);
        }
    }
    (row, col)
}

pub fn build_display_plan(ranges: &[MergeRange]) -> DisplayPlan {
    let mut skip = HashSet::new();
    let mut hskip_cols_by_row: HashMap<u32, HashSet<u32>> = HashMap::new();

    for rg in ranges {
        let is_horizontal = rg.min_col != rg.max_col;
        for r in rg.min_row..=rg.max_row {
            for c in rg.min_col..=rg.max_col {
                if (r, c) != (rg.min_row, rg.min_col) {
                    skip.insert((r, c));
                }
                if is_horizontal && c != rg.min_col {
                    hskip_cols_by_row.entry(r).or_default().insert(c);
                }
            }
        }
    }

    DisplayPlan {
        skip,
        hskip_cols_by_row,
    }
}

/// Với mỗi dòng, danh sách cột "thực" còn hiển thị (loại các cột bị merge
/// ngang nuốt) — tương đương `col_map_per_row` trong `compute_mergeinfo()`.
pub fn col_map_per_row(plan: &DisplayPlan, max_row: u32, max_col: u32) -> Vec<Vec<u32>> {
    (1..=max_row)
        .map(|r| {
            let drop = plan.hskip_cols_by_row.get(&r);
            (1..=max_col)
                .filter(|c| drop.map(|d| !d.contains(c)).unwrap_or(true))
                .collect()
        })
        .collect()
}

/// Giống `build_display_rows()` (đã gộp vào `build_display_rows_with_coords`)
/// toạ độ (row, col) GỐC trong sheet của từng cell hiển thị — cần để tra
/// style (CellStyle) đúng ô khi xuất metadata highlight cho Vim. Với cell bị
/// `skip` (nằm trong vùng merge nhưng không phải top-left), style vẫn nên
/// lấy từ chính top-left của vùng merge đó — main.rs xử lý việc này khi
/// tổng hợp metadata (vì plan.skip không tự nói top-left là ô nào).
pub fn build_display_rows_with_coords(
    sheet: &SheetData,
) -> (Vec<Vec<String>>, Vec<Vec<(u32, u32)>>) {
    let plan = build_display_plan(&sheet.merges);
    let max_row = sheet.max_row.max(1);
    let max_col = sheet.max_col.max(1);

    let mut display_rows = Vec::with_capacity(max_row as usize);
    let mut coords_rows = Vec::with_capacity(max_row as usize);

    for r in 1..=max_row {
        let cols_to_drop = plan.hskip_cols_by_row.get(&r);
        let mut display_row = Vec::new();
        let mut coords_row = Vec::new();
        for c in 1..=max_col {
            if cols_to_drop.map(|d| d.contains(&c)).unwrap_or(false) {
                continue;
            }
            coords_row.push((r, c));
            if plan.skip.contains(&(r, c)) {
                display_row.push(String::new());
            } else {
                let v = sheet.get(r, c).unwrap_or("");
                let escaped = v.replace('\r', " ").replace('\n', " \\n ");
                display_row.push(escaped);
            }
        }
        display_rows.push(display_row);
        coords_rows.push(coords_row);
    }

    if display_rows.is_empty() {
        display_rows.push(vec![String::new()]);
        coords_rows.push(vec![(1, 1)]);
    }

    (display_rows, coords_rows)
}

/// Kiểm tra (sheet_row, sheet_col) có thuộc cùng 1 vùng merge với
/// (sheet_row+1, sheet_col) không — tức ranh giới ngang giữa 2 dòng này bị
/// "nuốt" bởi merge dọc, không nên vẽ vạch '---' khi render bảng ASCII.
pub fn is_vmerge_below(ranges: &[MergeRange], row: u32, col: u32) -> bool {
    for rg in ranges {
        if row >= rg.min_row
            && row < rg.max_row
            && col >= rg.min_col
            && col <= rg.max_col
        {
            return true;
        }
    }
    false
}

/// Giống `build_display_rows_with_coords` nhưng trả thêm 2 ma trận:
/// - `vmerge_below[i][j] = true` nghĩa cell (i, j) merge dọc với (i+1, j),
///   dùng để xoá vạch ngang giữa 2 dòng khi render.
/// - `col_spans[i][j]` = số cột sheet mà cell hiển thị (i, j) đang chiếm.
///   Cell không merge ngang -> col_span = 1; cell merge ngang -> col_span =
///   max_col - min_col + 1 của vùng merge. Dùng để render width của cell
///   merge bằng tổng widths của các cột bị nuốt (giữ alignment ngang).
pub fn build_display_rows_full(
    sheet: &SheetData,
) -> (
    Vec<Vec<String>>,
    Vec<Vec<(u32, u32)>>,
    Vec<Vec<bool>>,
    Vec<Vec<u32>>,
) {
    let (display_rows, coords_rows) = build_display_rows_with_coords(sheet);

    let mut vmerge_below: Vec<Vec<bool>> = Vec::with_capacity(display_rows.len());
    let mut col_spans: Vec<Vec<u32>> = Vec::with_capacity(display_rows.len());

    for coords_row in &coords_rows {
        let row_vmerge: Vec<bool> = coords_row
            .iter()
            .map(|&(r, c)| is_vmerge_below(&sheet.merges, r, c))
            .collect();
        let row_spans: Vec<u32> = coords_row
            .iter()
            .map(|&(r, c)| {
                // Tìm vùng merge có top-left = (r, c). Nếu không phải top-left
                // của merge nào -> col_span = 1.
                for rg in &sheet.merges {
                    if rg.min_row == r && rg.min_col == c {
                        return rg.max_col - rg.min_col + 1;
                    }
                    // Cell ở giữa vùng merge dọc (nhưng top-left của col):
                    // không phải top-left của hàng, nhưng vẫn là "ô đại diện"
                    // của cột này — col_span dựa theo bề rộng cột của merge.
                    if r > rg.min_row
                        && r <= rg.max_row
                        && c == rg.min_col
                        && rg.min_col != rg.max_col
                    {
                        return rg.max_col - rg.min_col + 1;
                    }
                }
                1
            })
            .collect();
        vmerge_below.push(row_vmerge);
        col_spans.push(row_spans);
    }
    (display_rows, coords_rows, vmerge_below, col_spans)
}
