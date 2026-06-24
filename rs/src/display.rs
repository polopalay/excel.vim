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

/// Sinh ra các dòng hiển thị (Vec<Vec<String>>) dùng cho `render()`,
/// tương đương phần thân vòng lặp trong `open_xlsx()` của Python.
pub fn build_display_rows(sheet: &SheetData) -> Vec<Vec<String>> {
    let plan = build_display_plan(&sheet.merges);
    let max_row = sheet.max_row.max(1);
    let max_col = sheet.max_col.max(1);

    let mut display_rows = Vec::with_capacity(max_row as usize);

    for r in 1..=max_row {
        let cols_to_drop = plan.hskip_cols_by_row.get(&r);
        let mut display_row = Vec::new();
        for c in 1..=max_col {
            if cols_to_drop.map(|d| d.contains(&c)).unwrap_or(false) {
                continue;
            }
            if plan.skip.contains(&(r, c)) {
                display_row.push(String::new());
            } else {
                let v = sheet.get(r, c).unwrap_or("");
                // Escape giống Python: \r -> ' ', \n -> ' \n ' để mỗi cell
                // luôn nằm trên đúng 1 dòng text trong bảng ASCII.
                let escaped = v.replace('\r', " ").replace('\n', " \\n ");
                display_row.push(escaped);
            }
        }
        display_rows.push(display_row);
    }

    if display_rows.is_empty() {
        display_rows.push(vec![String::new()]);
    }

    display_rows
}
