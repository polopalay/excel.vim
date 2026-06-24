//! Render & parse bảng ASCII — port 1:1 từ `render()` / `parse_ascii()`
//! trong excel.py. Đây cũng chính là vòng lặp Python "tệ" mà anh muốn
//! chuyển sang Rust để nhanh hơn trên file lớn.

/// Chuyển 1 danh sách dòng (mỗi dòng là Vec<String>) thành bảng ASCII có
/// viền `+---+---+` và dấu `|` phân cách cột, căn trái theo cột rộng nhất.
pub fn render(rows: &[Vec<String>]) -> String {
    if rows.is_empty() {
        return String::new();
    }

    let ncols = rows.iter().map(|r| r.len()).max().unwrap_or(0);
    let mut widths = vec![0usize; ncols];
    for row in rows {
        for (i, v) in row.iter().enumerate() {
            // Dùng số ký tự Unicode (chars().count()), không phải bytes,
            // để căn cột đúng với chuỗi có dấu tiếng Việt (khớp hành vi
            // len() của Python trên str, vốn đếm theo code point).
            let len = v.chars().count();
            if len > widths[i] {
                widths[i] = len;
            }
        }
    }

    let border: String = {
        let mut s = String::from("+");
        for w in &widths {
            s.push_str(&"-".repeat(w + 2));
            s.push('+');
        }
        s
    };

    let mut out = String::with_capacity(border.len() * (rows.len() * 2 + 1));
    out.push_str(&border);
    out.push('\n');

    for row in rows {
        out.push('|');
        for (i, v) in row.iter().enumerate() {
            out.push(' ');
            out.push_str(v);
            let pad = widths[i].saturating_sub(v.chars().count());
            out.push_str(&" ".repeat(pad));
            out.push_str(" |");
        }
        out.push('\n');
        out.push_str(&border);
        out.push('\n');
    }

    // Bỏ newline cuối cùng để khớp "\n".join(out) của Python (không có
    // newline thừa ở cuối chuỗi).
    if out.ends_with('\n') {
        out.pop();
    }
    out
}

/// Ngược lại render(): từ các dòng text bảng ASCII -> Vec<Vec<String>>.
/// Chỉ dòng bắt đầu bằng '|' mới được xử lý; dòng viền '+' bị bỏ qua.
pub fn parse_ascii(lines: &[String]) -> Vec<Vec<String>> {
    let mut rows = Vec::new();
    for line in lines {
        let line = line.trim_end_matches(['\r', '\n']);
        if !line.starts_with('|') {
            continue;
        }
        let parts: Vec<&str> = line.split('|').collect();
        // split('|') trên dòng "|a|b|" cho ["", "a", "b", ""] -> bỏ đầu/cuối
        if parts.len() < 2 {
            rows.push(Vec::new());
            continue;
        }
        let cols: Vec<String> = parts[1..parts.len() - 1]
            .iter()
            .map(|c| c.trim().to_string())
            .collect();
        rows.push(cols);
    }
    rows
}
