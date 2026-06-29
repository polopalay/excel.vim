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

/// Màu RGB (không có alpha riêng — luôn coi như FF + RGB khi ghi ra XML,
/// khớp với cách Excel/openpyxl thường ghi "FFRRGGBB").
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct Rgb {
    pub r: u8,
    pub g: u8,
    pub b: u8,
}

impl Rgb {
    pub fn to_hex(&self) -> String {
        format!("{:02X}{:02X}{:02X}", self.r, self.g, self.b)
    }

    /// Parse từ "RRGGBB" hoặc "AARRGGBB" (8 hex digit, bỏ 2 ký tự alpha đầu).
    pub fn from_hex(s: &str) -> Option<Rgb> {
        let s = s.trim_start_matches('#');
        let hex = if s.len() == 8 { &s[2..] } else if s.len() == 6 { s } else { return None };
        let r = u8::from_str_radix(&hex[0..2], 16).ok()?;
        let g = u8::from_str_radix(&hex[2..4], 16).ok()?;
        let b = u8::from_str_radix(&hex[4..6], 16).ok()?;
        Some(Rgb { r, g, b })
    }
}

/// Style đã "resolve" của 1 cell: đủ thông tin để vừa hiển thị (Vim syntax
/// highlight) vừa ghi lại đúng vào styles.xml khi save.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct CellStyle {
    pub bold: bool,
    pub italic: bool,
    pub font_color: Option<Rgb>,
    pub bg_color: Option<Rgb>,
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
    /// style_id GỐC (attribute s="N" trong file .xlsx) của từng cell có
    /// style khác mặc định. Giữ lại để khi ghi ra XML ta tái sử dụng đúng
    /// style cũ, không làm mất style khi user chỉ sửa nội dung text.
    pub cell_style_id: BTreeMap<(u32, u32), u32>,
    /// Style đã resolve (bold/italic/màu) tương ứng với từng style_id,
    /// dùng để hiển thị (xuất metadata cho Vim) và để tính style mới khi
    /// user đổi màu nền qua lệnh setbg.
    pub styles: Vec<CellStyle>,
    /// Công thức GỐC (không có dấu "=" ở đầu) của các cell có thẻ <f> trong
    /// XML, ví dụ {(5,2): "SUM(A1:A4)"}. `cells` vẫn lưu giá trị HIỂN THỊ
    /// (kết quả đã tính) — display.rs/table.rs không cần biết gì về
    /// formula cả, chúng chỉ đọc `cells` như bình thường.
    pub formulas: BTreeMap<(u32, u32), String>,
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

    /// Lấy CellStyle đã resolve của 1 cell (style mặc định nếu cell không
    /// có style riêng / style_id ngoài phạm vi `styles`).
    pub fn get_style(&self, row: u32, col: u32) -> CellStyle {
        self.cell_style_id
            .get(&(row, col))
            .and_then(|&id| self.styles.get(id as usize))
            .cloned()
            .unwrap_or_default()
    }

    /// Tìm style_id đã có sẵn khớp đúng CellStyle này, hoặc thêm mới vào
    /// `styles` và trả về index mới. Dùng khi user đổi màu nền/style của 1
    /// cell — tránh phình to styles.xml nếu nhiều cell dùng chung 1 style.
    fn intern_style(&mut self, style: CellStyle) -> u32 {
        if let Some(idx) = self.styles.iter().position(|s| *s == style) {
            return idx as u32;
        }
        self.styles.push(style);
        (self.styles.len() - 1) as u32
    }

    /// Mở rộng max_row/max_col để bao gồm (row, col) nếu nó đang nằm ngoài
    /// phạm vi hiện tại — dùng khi style được áp cho 1 cell trống ngoài
    /// vùng dữ liệu hiện có (ví dụ tô màu nền cho ô D1 khi sheet mới chỉ có
    /// tới cột C), để cell đó thực sự xuất hiện trong bảng hiển thị.
    fn expand_bounds(&mut self, row: u32, col: u32) {
        if row > self.max_row {
            self.max_row = row;
        }
        if col > self.max_col {
            self.max_col = col;
        }
    }

    /// Đặt màu nền cho 1 cell, giữ nguyên bold/italic/font_color hiện tại.
    /// Truyền `None` để xoá màu nền (về lại mặc định "không tô màu").
    pub fn set_bg_color(&mut self, row: u32, col: u32, bg: Option<Rgb>) {
        self.expand_bounds(row, col);
        let mut style = self.get_style(row, col);
        style.bg_color = bg;
        let new_id = self.intern_style(style);
        self.cell_style_id.insert((row, col), new_id);
    }

    /// Đặt màu chữ cho 1 cell, giữ nguyên bold/italic/bg_color hiện tại.
    /// Truyền `None` để xoá màu chữ (về lại mặc định "màu tự động/đen").
    pub fn set_fg_color(&mut self, row: u32, col: u32, fg: Option<Rgb>) {
        self.expand_bounds(row, col);
        let mut style = self.get_style(row, col);
        style.font_color = fg;
        let new_id = self.intern_style(style);
        self.cell_style_id.insert((row, col), new_id);
    }

    /// Đảo trạng thái bold của 1 cell (bật nếu đang tắt, tắt nếu đang bật),
    /// giữ nguyên italic/màu chữ/màu nền hiện tại.
    pub fn toggle_bold(&mut self, row: u32, col: u32) {
        self.expand_bounds(row, col);
        let mut style = self.get_style(row, col);
        style.bold = !style.bold;
        let new_id = self.intern_style(style);
        self.cell_style_id.insert((row, col), new_id);
    }

    /// Đảo trạng thái italic của 1 cell, giữ nguyên bold/màu chữ/màu nền.
    pub fn toggle_italic(&mut self, row: u32, col: u32) {
        self.expand_bounds(row, col);
        let mut style = self.get_style(row, col);
        style.italic = !style.italic;
        let new_id = self.intern_style(style);
        self.cell_style_id.insert((row, col), new_id);
    }

    /// Gộp 1 vùng (min_row..=max_row, min_col..=max_col) thành 1 ô — tương
    /// đương Excel's "Merge Cells". Tuân theo đúng hành vi Excel:
    ///
    /// 1. Mọi merge cũ giao với vùng mới sẽ bị xoá (Excel: thay merge cũ
    ///    bằng merge mới chồng lên).
    /// 2. Giá trị TEXT của các cell không phải top-left trong vùng mới sẽ
    ///    bị xoá — chỉ giữ value của ô top-left (giống Excel).
    /// 3. Style của các cell không phải top-left vẫn được giữ trong XML
    ///    (Excel cũng giữ — vì user có thể bỏ gộp sau và muốn style cũ).
    ///
    /// Trả về `true` nếu thực sự thêm được merge, `false` nếu vùng chỉ có
    /// 1 cell (không thực sự gộp gì) hoặc bounds không hợp lệ.
    pub fn add_merge(&mut self, min_row: u32, min_col: u32, max_row: u32, max_col: u32) -> bool {
        if min_row > max_row || min_col > max_col {
            return false;
        }
        if min_row == max_row && min_col == max_col {
            return false; // 1 cell -> không gộp gì
        }

        // Xoá mọi merge cũ giao với vùng mới
        self.merges.retain(|m| {
            !(m.min_row <= max_row
                && m.max_row >= min_row
                && m.min_col <= max_col
                && m.max_col >= min_col)
        });

        // Xoá value các cell không phải top-left trong vùng mới
        for r in min_row..=max_row {
            for c in min_col..=max_col {
                if (r, c) != (min_row, min_col) {
                    self.cells.remove(&(r, c));
                }
            }
        }

        self.merges.push(MergeRange {
            min_row,
            max_row,
            min_col,
            max_col,
        });
        self.expand_bounds(max_row, max_col);
        true
    }

    /// Bỏ gộp mọi vùng merge giao với 1 cell hoặc 1 range cho trước. Trả về
    /// số lượng merge đã bỏ.
    pub fn remove_merges_intersecting(
        &mut self,
        min_row: u32,
        min_col: u32,
        max_row: u32,
        max_col: u32,
    ) -> usize {
        let before = self.merges.len();
        self.merges.retain(|m| {
            !(m.min_row <= max_row
                && m.max_row >= min_row
                && m.min_col <= max_col
                && m.max_col >= min_col)
        });
        before - self.merges.len()
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
