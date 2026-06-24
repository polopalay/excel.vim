//! Mô hình dữ liệu workbook trong bộ nhớ.
//!
//! Khác với openpyxl (load toàn bộ object model phức tạp: styles, themes,
//! charts...), ở đây ta chỉ giữ đủ thông tin mà excel.py gốc thực sự dùng:
//! - giá trị từng cell (dạng String, đã resolve qua sharedStrings)
//! - danh sách vùng merge
//! - kích thước max_row/max_col
//! - danh sách tên sheet + sheet nào đang active
//!
//! Mọi phần khác của file .xlsx (styles.xml, theme1.xml, các sheet khác,
//! drawings, ...) được giữ nguyên bytes-for-bytes khi ghi lại, để không làm
//! hỏng định dạng / hình ảnh / công thức của các sheet không bị đụng tới.

use std::collections::BTreeMap;

#[derive(Debug, Clone, Default)]
pub struct MergeRange {
    pub min_row: u32,
    pub max_row: u32,
    pub min_col: u32,
    pub max_col: u32,
}

/// Một worksheet đã được parse: map (row, col) -> giá trị string hiển thị.
/// Dùng BTreeMap để duyệt theo thứ tự ổn định khi cần (không bắt buộc,
/// nhưng giúp debug dễ hơn HashMap).
#[derive(Debug, Clone, Default)]
pub struct SheetData {
    pub cells: BTreeMap<(u32, u32), String>,
    pub merges: Vec<MergeRange>,
    pub max_row: u32,
    pub max_col: u32,
}

impl SheetData {
    pub fn get(&self, row: u32, col: u32) -> Option<&str> {
        self.cells.get(&(row, col)).map(|s| s.as_str())
    }

    pub fn set(&mut self, row: u32, col: u32, value: String) {
        if value.is_empty() {
            self.cells.remove(&(row, col));
        } else {
            self.cells.insert((row, col), value);
        }
        if row > self.max_row {
            self.max_row = row;
        }
        if col > self.max_col {
            self.max_col = col;
        }
    }

    /// Mở rộng (không bao giờ co lại) max_row/max_col dựa trên dữ liệu cell
    /// thực tế và các vùng merge. Gọi sau khi đã set max_row/max_col từ
    /// <dimension> hoặc từ các cell rỗng cuối bảng, để không bị mất thông
    /// tin kích thước sheet khi cột/dòng cuối không có giá trị nào.
    pub fn recompute_bounds(&mut self) {
        let mut max_row = self.max_row;
        let mut max_col = self.max_col;
        for &(r, c) in self.cells.keys() {
            if r > max_row {
                max_row = r;
            }
            if c > max_col {
                max_col = c;
            }
        }
        for m in &self.merges {
            if m.max_row > max_row {
                max_row = m.max_row;
            }
            if m.max_col > max_col {
                max_col = m.max_col;
            }
        }
        self.max_row = max_row.max(1);
        self.max_col = max_col.max(1);
    }
}

/// Thông tin 1 sheet trong workbook.xml: tên hiển thị + đường dẫn file XML
/// thực tế bên trong zip (xl/worksheets/sheetN.xml), vì openpyxl/Excel có
/// thể đặt tên sheet khác thứ tự file vật lý.
#[derive(Debug, Clone)]
pub struct SheetEntry {
    pub name: String,
    /// Đường dẫn trong zip, ví dụ "xl/worksheets/sheet1.xml"
    pub path: String,
    /// r:id trong workbook.xml.rels, ví dụ "rId1"
    pub rid: String,
    /// sheetId trong workbook.xml (số nguyên dùng để khôi phục)
    pub sheet_id: u32,
}
