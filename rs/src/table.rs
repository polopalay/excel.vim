//! Render & parse bảng ASCII — port 1:1 từ `render()` / `parse_ascii()`
//! trong excel.py. Đây cũng chính là vòng lặp Python "tệ" mà anh muốn
//! chuyển sang Rust để nhanh hơn trên file lớn.

use std::collections::HashSet;

/// Render bảng ASCII có hỗ trợ merge dọc + merge ngang. Widths được tính
/// THEO CỘT SHEET (không phải index trong row), bằng cách chỉ dùng các
/// cell `col_span == 1` (không bị merge ngang) — vì cell merge ngang
/// "đáng lẽ" chiếm chỗ của nhiều cột, không nên ảnh hưởng width 1 cột riêng.
///
/// Cell merge ngang khi render sẽ chiếm chỗ = tổng widths của các cột nó
/// nuốt + 3 ký tự " | " cho mỗi border bị nuốt giữa các cột (giữ alignment
/// với các dòng khác không bị merge).
///
/// Tham số:
///   - rows[i][j]: text của cell hiển thị thứ j ở row i.
///   - coords_rows[i][j]: (sheet_row, sheet_col) của cell đó. Dùng để biết
///     cell tương ứng cột sheet nào.
///   - col_spans[i][j]: số cột sheet mà cell đó chiếm (1 nếu không merge
///     ngang). Phải có cùng shape với rows.
///   - vmerge_below[i][j]: cell (i, j) merge dọc với (i+1, j) -> border
///     dưới bị "ăn".
///
/// `vmerge_below` và `col_spans` có thể là slice rỗng (hành vi cũ: không
/// merge gì cả, mọi cell col_span = 1, border đầy đủ).
///
/// Trả về (text_output, spans) — spans[i][j] = (char_start, char_end) 1-based
/// inclusive-start exclusive-end của cell j trong row i (cho Vim metadata).
pub fn render_with_spans(
    rows: &[Vec<String>],
    coords_rows: &[Vec<(u32, u32)>],
    col_spans: &[Vec<u32>],
    vmerge_below: &[Vec<bool>],
) -> (
    String,
    Vec<Vec<(usize, usize)>>,
    Vec<Vec<(usize, usize)>>,
) {
    if rows.is_empty() {
        return (String::new(), Vec::new(), Vec::new());
    }

    // Bước 1: tìm tập cột sheet thực sự xuất hiện trong bảng. Mỗi cột sheet
    // được "đại diện" bởi cell có col_span = 1 chiếm nó. Cell merge ngang
    // tại cột min_col còn các cột min_col+1..=max_col bị skip trên hàng đó.
    let max_sheet_col = coords_rows
        .iter()
        .flatten()
        .map(|&(_, c)| c)
        .max()
        .unwrap_or(1);

    // Bước 2: tính width của TỪNG cột sheet (1..=max_sheet_col). Chỉ xét
    // các cell col_span = 1 (không bị merge ngang) — vì cell merge ngang
    // "chiếm chỗ" của nhiều cột nên không phản ánh width 1 cột riêng lẻ.
    let mut col_widths: Vec<usize> = vec![0; (max_sheet_col + 1) as usize];
    for (row_idx, row) in rows.iter().enumerate() {
        for (cell_idx, v) in row.iter().enumerate() {
            let span = col_spans
                .get(row_idx)
                .and_then(|s| s.get(cell_idx))
                .copied()
                .unwrap_or(1);
            if span != 1 {
                continue;
            }
            let (_, sc) = coords_rows[row_idx][cell_idx];
            let len = v.chars().count();
            if len > col_widths[sc as usize] {
                col_widths[sc as usize] = len;
            }
        }
    }
    // Đảm bảo mọi cột đều có width tối thiểu nào đó để render đúng — nếu
    // 1 cột chỉ xuất hiện qua cell merge ngang (không có cell single-span
    // nào trên cột đó), gán width tối thiểu = 1.
    for w in col_widths.iter_mut().skip(1) {
        if *w == 0 {
            *w = 1;
        }
    }

    // Bước 2.5: phân phối lại width nếu cell MERGE NGANG có nội dung dài
    // hơn tổng widths của các cột thành phần. Công thức (theo yêu cầu):
    //   số_dư = len(content) - sum(widths cột thành phần) - 3*(span-1)
    //   trung_bình = số_dư / span    (chia đều)
    //   số_lẻ = số_dư % span         (cộng vào CỘT ĐẦU của vùng merge)
    //
    // Cách này đảm bảo cell merge ngang có đủ chỗ cho content, đồng thời
    // các cột thành phần đều được "nới" tương ứng -> mọi dòng KHÁC trong
    // cùng cột (không bị merge ngang) vẫn căn ngay ngắn với cell merged,
    // không bị bảng "gãy" alignment.
    //
    // Lặp nhiều pass cho đến khi không còn cell merge nào "tràn" — vì việc
    // nới 1 cột để chứa cell merge có thể vẫn không đủ cho 1 cell merge
    // khác (chồng lấp), hoặc ngược lại. Trong thực tế hội tụ rất nhanh
    // (1-2 pass) vì mỗi pass chỉ phân phối phần dư còn thiếu.
    loop {
        let mut changed = false;
        for (row_idx, row) in rows.iter().enumerate() {
            for (cell_idx, v) in row.iter().enumerate() {
                let span = col_spans
                    .get(row_idx)
                    .and_then(|s| s.get(cell_idx))
                    .copied()
                    .unwrap_or(1);
                if span <= 1 {
                    continue;
                }
                let (_, sc) = coords_rows[row_idx][cell_idx];
                let len = v.chars().count();
                let mut current_total: usize = 0;
                for cc in sc..(sc + span) {
                    current_total += col_widths[cc as usize];
                }
                // Tổng "chỗ" thực tế cell merge có = sum(widths) + 3*(span-1)
                let available = current_total + 3 * (span as usize - 1);
                if len <= available {
                    continue;
                }
                let extra = len - available;
                let avg = extra / span as usize;
                let rem = extra % span as usize;
                // Cộng `avg` vào mọi cột thành phần, cộng thêm `rem` vào cột
                // đầu (theo yêu cầu).
                for (k, cc) in (sc..(sc + span)).enumerate() {
                    let add = if k == 0 { avg + rem } else { avg };
                    if add > 0 {
                        col_widths[cc as usize] += add;
                        changed = true;
                    }
                }
            }
        }
        if !changed {
            break;
        }
    }

    // Bước 3: với mỗi cell merge ngang, tính width hiển thị = sum widths
    // cột nó chiếm + 3 * (span - 1). Lưu trong cell_display_widths[i][j].
    let mut cell_display_widths: Vec<Vec<usize>> = Vec::with_capacity(rows.len());
    for (row_idx, row) in rows.iter().enumerate() {
        let mut row_widths = Vec::with_capacity(row.len());
        for (cell_idx, v) in row.iter().enumerate() {
            let span = col_spans
                .get(row_idx)
                .and_then(|s| s.get(cell_idx))
                .copied()
                .unwrap_or(1);
            let (_, sc) = coords_rows[row_idx][cell_idx];
            let mut base = 0usize;
            for cc in sc..(sc + span) {
                base += col_widths[cc as usize];
            }
            // Mỗi border bị nuốt giữa các cột = 3 ký tự " | "
            base += 3 * (span as usize - 1);
            // Sau pass phân phối, base luôn >= len. Nhưng vẫn safety check
            // cho cell single-span có content dài bất thường (đã được lo
            // ở pass tính col_widths, nhưng để chắc).
            let len = v.chars().count();
            if len > base {
                base = len;
            }
            row_widths.push(base);
        }
        cell_display_widths.push(row_widths);
    }

    // Bước 4: dựng full border (dòng trên cùng và dưới cùng) — gồm '+' và
    // '-' theo widths của TỪNG CỘT SHEET (luôn đầy đủ, không nhóm theo merge).
    let full_border: String = {
        let mut s = String::from("+");
        for c in 1..=max_sheet_col {
            s.push_str(&"-".repeat(col_widths[c as usize] + 2));
            s.push('+');
        }
        s
    };

    let mut out = String::with_capacity(full_border.len() * (rows.len() * 2 + 1));
    out.push_str(&full_border);
    out.push('\n');

    let mut all_spans = Vec::with_capacity(rows.len());
    let mut all_full_spans = Vec::with_capacity(rows.len());

    // Bước 5: render từng content row + border dưới.
    for (row_idx, row) in rows.iter().enumerate() {
        let mut col_pos = 0usize; // 0-based char index
        let mut spans = Vec::with_capacity(row.len());
        let mut full_spans = Vec::with_capacity(row.len());

        out.push('|');
        col_pos += 1;
        for (cell_idx, v) in row.iter().enumerate() {
            // Bắt đầu của TOÀN cell (kể cả padding " " trước text). Đây
            // là vị trí ngay sau vạch '|' bên trái cell.
            let full_start = col_pos + 1;
            out.push(' ');
            col_pos += 1;
            let start = col_pos + 1;
            out.push_str(v);
            let len = v.chars().count();
            col_pos += len;
            let end = col_pos + 1;
            spans.push((start, end));

            let cell_w = cell_display_widths[row_idx][cell_idx];
            let pad = cell_w.saturating_sub(len);
            out.push_str(&" ".repeat(pad));
            col_pos += pad;
            // Kết thúc của TOÀN cell (kể cả padding " " sau text). Đây là
            // vị trí ngay TRƯỚC vạch '|' bên phải cell.
            out.push(' ');
            col_pos += 1;
            let full_end = col_pos + 1;
            full_spans.push((full_start, full_end));
            out.push('|');
            col_pos += 1;
        }
        out.push('\n');

        // Border dưới: full nếu là dòng cuối, hoặc tuỳ chỉnh theo
        // vmerge_below và col_spans của row tiếp theo (để biết những vạch
        // dọc nào bị "ăn" do cell trên hoặc cell dưới merge ngang qua đó).
        let is_last = row_idx + 1 == rows.len();
        if is_last {
            out.push_str(&full_border);
        } else {
            let empty_b: Vec<bool> = Vec::new();
            let vmerge_row = vmerge_below.get(row_idx).unwrap_or(&empty_b);
            // Tính tập cột sheet bị "ăn" border ngang giữa row_idx và
            // row_idx+1: cột sheet c bị ăn nếu cell tại (row_idx, c) merge
            // dọc xuống cell dưới (vmerge_below true cho cột chứa c).
            let mut vmerged_cols: HashSet<u32> = HashSet::new();
            for (cell_idx, &(_, sc)) in coords_rows[row_idx].iter().enumerate() {
                if vmerge_row.get(cell_idx).copied().unwrap_or(false) {
                    let span = col_spans
                        .get(row_idx)
                        .and_then(|s| s.get(cell_idx))
                        .copied()
                        .unwrap_or(1);
                    for c in sc..(sc + span) {
                        vmerged_cols.insert(c);
                    }
                }
            }

            // Tính tập cột bị merge ngang TRÊN (row_idx) và DƯỚI (row_idx+1):
            // 1 vạch dọc giữa cột c và c+1 bị "ăn" trên row đó nếu CẢ
            // (row, c) lẫn (row, c+1) đều thuộc cùng 1 cell merge ngang
            // (= không có border giữa 2 cột này trên content row).
            let hmerged_top = collect_hmerged_borders(coords_rows.get(row_idx).unwrap_or(&Vec::new()), col_spans.get(row_idx).unwrap_or(&Vec::new()));
            let hmerged_bot = collect_hmerged_borders(coords_rows.get(row_idx + 1).unwrap_or(&Vec::new()), col_spans.get(row_idx + 1).unwrap_or(&Vec::new()));

            out.push_str(&build_partial_border_full(
                &col_widths,
                max_sheet_col,
                &vmerged_cols,
                &hmerged_top,
                &hmerged_bot,
            ));
        }
        out.push('\n');

        all_spans.push(spans);
        all_full_spans.push(full_spans);
    }

    if out.ends_with('\n') {
        out.pop();
    }

    (out, all_spans, all_full_spans)
}

/// Trả về tập các vị trí "giữa 2 cột" (giá trị c nghĩa là giữa cột c và c+1)
/// mà 1 cell merge ngang đang vắt qua trên 1 row cụ thể — tức vị trí KHÔNG
/// có vạch dọc `|` trên row đó vì nó nằm bên trong 1 cell merge.
fn collect_hmerged_borders(
    coords_row: &[(u32, u32)],
    spans_row: &[u32],
) -> HashSet<u32> {
    let mut set = HashSet::new();
    for (i, &(_, sc)) in coords_row.iter().enumerate() {
        let span = spans_row.get(i).copied().unwrap_or(1);
        if span > 1 {
            // Cell chiếm cột [sc, sc+span-1]; vạch dọc bị ăn ở vị trí
            // sc, sc+1, ..., sc+span-2 (giữa các cột trong vùng merge).
            for c in sc..(sc + span - 1) {
                set.insert(c);
            }
        }
    }
    set
}

/// Dựng partial border giữa 2 content row, có hỗ trợ:
///   - `vmerged_cols`: tập cột bị merge dọc qua border này -> thay '---'
///     bằng spaces tại cột đó.
///   - `hmerged_top` / `hmerged_bot`: tập vị trí "giữa cột c và c+1" mà
///     cell merge ngang đang vắt qua trên row trên / dưới -> không vẽ '+'
///     tại vị trí đó (thay bằng ' ' hoặc '-' tuỳ context).
fn build_partial_border_full(
    col_widths: &[usize],
    max_sheet_col: u32,
    vmerged_cols: &HashSet<u32>,
    hmerged_top: &HashSet<u32>,
    hmerged_bot: &HashSet<u32>,
) -> String {
    let mut s = String::new();

    // Junction tại cột c (1..=max_sheet_col+1): vị trí của ký tự '+' chia
    // giữa col c-1 (bên trái) và c (bên phải) trong border.
    // - Junction đầu tiên (c=1): bên trái không có cột, bên phải là cột 1.
    // - Junction cuối (c=max_sheet_col+1): bên trái là cột cuối, bên phải
    //   không có cột.
    for c in 1..=max_sheet_col {
        // Junction bên trái cell c (giữa cột c-1 và c, hoặc đầu bảng nếu c=1)
        let junction_char = compute_junction(c, max_sheet_col, vmerged_cols, hmerged_top, hmerged_bot);
        s.push(junction_char);

        // Phần border ngang cho cell c: '-' nếu cột c KHÔNG bị merge dọc,
        // ' ' nếu BỊ merge dọc.
        let fill = if vmerged_cols.contains(&c) { ' ' } else { '-' };
        s.push_str(&fill.to_string().repeat(col_widths[c as usize] + 2));
    }
    // Junction cuối cùng (sau cột cuối): luôn là góc phải, không có cột bên phải.
    let last_j = compute_junction(max_sheet_col + 1, max_sheet_col, vmerged_cols, hmerged_top, hmerged_bot);
    s.push(last_j);
    s
}

/// Tính ký tự junction (chỗ giao 4 hướng) tại "khoảng giữa cột c-1 và c"
/// trên partial border. Có 4 trạng thái nhị phân:
///   - vạch ngang TRÁI (cột c-1 KHÔNG merge dọc) -> '-' tiếp giáp
///   - vạch ngang PHẢI (cột c KHÔNG merge dọc) -> '-' tiếp giáp
///   - vạch dọc TRÊN (không bị merge ngang bên trên ở vị trí này)
///   - vạch dọc DƯỚI (không bị merge ngang bên dưới ở vị trí này)
/// Junction = '+' nếu có ít nhất 1 ngang VÀ 1 dọc tiếp giáp; '|' nếu chỉ
/// có dọc (trên và/hoặc dưới); '-' nếu chỉ có ngang; ' ' nếu cô lập.
fn compute_junction(
    c: u32,
    max_sheet_col: u32,
    vmerged_cols: &HashSet<u32>,
    hmerged_top: &HashSet<u32>,
    hmerged_bot: &HashSet<u32>,
) -> char {
    let has_horiz_left = c > 1 && !vmerged_cols.contains(&(c - 1));
    let has_horiz_right = c <= max_sheet_col && !vmerged_cols.contains(&c);
    // Vạch dọc trên/dưới ở vị trí "giữa cột c-1 và c" tương ứng index c-1
    // trong hmerged_*: cell merge ngang ăn vạch dọc giữa cột (c-1) và c.
    // c = 1 hoặc c = max_sheet_col+1 nghĩa là biên trái/phải -> không có
    // merge ngang ăn được (vì không có cột bên ngoài bảng).
    let inside = c > 1 && c <= max_sheet_col;
    let key = c - 1; // index "giữa cột key và key+1"
    let has_vert_top = !inside || !hmerged_top.contains(&key);
    let has_vert_bot = !inside || !hmerged_bot.contains(&key);
    // Biên trái/phải của bảng: luôn có vạch dọc (viền ngoài), bất kể merge.
    let has_vert_top = has_vert_top || !inside;
    let has_vert_bot = has_vert_bot || !inside;

    let any_horiz = has_horiz_left || has_horiz_right;
    let any_vert = has_vert_top || has_vert_bot;
    match (any_horiz, any_vert) {
        (true, true) => '+',
        (false, true) => '|',
        (true, false) => '-',
        (false, false) => ' ',
    }
}

/// Parse bảng ASCII -> Vec<Vec<String>>.
///
/// Phân biệt dòng content và dòng border:
///   - Dòng border CHỈ chứa các ký tự '+', '-', '|', và space (kể cả partial
///     border của merge dọc, vd "|       +------+-------+      |").
///   - Dòng content luôn chứa ÍT NHẤT 1 ký tự khác 4 ký tự trên (ký tự nội
///     dung thực — chữ, số, dấu).
///
/// Quy tắc này có 1 trường hợp edge: cell content chỉ chứa space + "-+|"
/// thì sẽ bị nhầm là border. Nhưng vì:
///   - Cell hoàn toàn rỗng (chỉ space) -> hiển thị thành 1 dòng " |   | "
///     và parse_ascii vẫn xử lý đúng vì split('|') không cần biết nó là
///     border hay không (border thuần "+---+" thì split cho ["", "---",""]
///     và trim ra "---" — nhưng vì dòng border luôn được skip nguyên cả
///     dòng, edge case này không xảy ra).
///   - Trong thực tế, người dùng không thường chỉ gõ "-+|" vào 1 ô.
pub fn parse_ascii(lines: &[String]) -> Vec<Vec<String>> {
    let mut rows = Vec::new();
    for line in lines {
        let line = line.trim_end_matches(['\r', '\n']);
        if !line.starts_with('|') && !line.starts_with('+') {
            continue;
        }
        // Border row: chỉ gồm các ký tự '+', '-', '|', space VÀ phải chứa
        // ít nhất 1 ký tự '+' hoặc '-' (để phân biệt với 1 content row
        // toàn ô trống — vd "|       |       |" chỉ space và '|' giữa các
        // ô rỗng cũng có chữ thoả "chỉ +-| space" nhưng không phải border).
        let is_border = line
            .chars()
            .all(|c| c == '+' || c == '-' || c == '|' || c == ' ')
            && line.chars().any(|c| c == '+' || c == '-');
        if is_border {
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
