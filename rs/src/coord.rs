//! Chuyển đổi toạ độ cell giữa dạng "A1" (chữ-số) và (col, row) số nguyên.
//! Tương đương `openpyxl.utils.get_column_letter` /
//! `openpyxl.utils.coordinate_to_tuple` trong bản Python.

use crate::error::{AppError, AppResult};

/// Số cột (1-based) -> chữ cột Excel ("A", "B", ..., "Z", "AA", ...).
pub fn col_to_letters(mut col: u32) -> String {
    let mut s = Vec::new();
    while col > 0 {
        let rem = (col - 1) % 26;
        s.push((b'A' + rem as u8) as char);
        col = (col - 1) / 26;
    }
    s.iter().rev().collect()
}

/// Chữ cột Excel -> số cột (1-based).
pub fn letters_to_col(letters: &str) -> u32 {
    let mut col: u32 = 0;
    for c in letters.chars() {
        col = col * 26 + (c.to_ascii_uppercase() as u32 - 'A' as u32 + 1);
    }
    col
}

/// Tách 1 ref cell dạng "B12" -> ("B", 12). Không phụ thuộc dấu $ (tuyệt đối),
/// vì sheetN.xml luôn ghi ref không có $.
pub fn split_ref(cell_ref: &str) -> AppResult<(String, u32)> {
    let letters_end = cell_ref
        .find(|c: char| c.is_ascii_digit())
        .ok_or_else(|| AppError(format!("Invalid cell ref: {cell_ref}")))?;
    let letters = &cell_ref[..letters_end];
    let digits = &cell_ref[letters_end..];
    let row: u32 = digits
        .parse()
        .map_err(|_| AppError(format!("Invalid cell ref: {cell_ref}")))?;
    Ok((letters.to_string(), row))
}

/// "B12" -> (col=2, row=12)
pub fn ref_to_col_row(cell_ref: &str) -> AppResult<(u32, u32)> {
    let (letters, row) = split_ref(cell_ref)?;
    Ok((letters_to_col(&letters), row))
}

/// (col, row) -> "B12"
pub fn col_row_to_ref(col: u32, row: u32) -> String {
    format!("{}{}", col_to_letters(col), row)
}

/// "A1:C5" -> (min_col, min_row, max_col, max_row)
pub fn parse_range(range: &str) -> AppResult<(u32, u32, u32, u32)> {
    let mut parts = range.split(':');
    let start = parts
        .next()
        .ok_or_else(|| AppError(format!("Invalid range: {range}")))?;
    let end = parts.next().unwrap_or(start);
    let (c1, r1) = ref_to_col_row(start)?;
    let (c2, r2) = ref_to_col_row(end)?;
    Ok((c1.min(c2), r1.min(r2), c1.max(c2), r1.max(r2)))
}
