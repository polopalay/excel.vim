import sys
import warnings
import os
from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter

warnings.filterwarnings("ignore", message="Data Validation extension is not supported*")


def ensure_workbook(path):
    if os.path.exists(path):
        return
    wb = Workbook()
    ws = wb.active
    ws["A1"] = ""
    wb.save(path)

def get_merge_ranges(ws):
    """Return all merged ranges as dictionaries containing row/column boundaries."""
    ranges = []
    for mc in ws.merged_cells.ranges:
        ranges.append(
            {
                "min_row": mc.min_row,
                "max_row": mc.max_row,
                "min_col": mc.min_col,
                "max_col": mc.max_col,
            }
        )
    return ranges

def build_display_plan(ranges, max_row, max_col):
    """
    Build the worksheet display plan.

    Returns:
        skip:
            Set of merged cells that should not display values.

        hskip_cols_by_row:
            Columns that should be completely hidden on a row because
            they are covered by a horizontal merge range.
    """
    skip = set()
    hskip_cols_by_row = {}
    top_left = {}
    for rg in ranges:
        top_left[(rg["min_row"], rg["min_col"])] = rg
        is_horizontal = rg["min_col"] != rg["max_col"]
        for r in range(rg["min_row"], rg["max_row"] + 1):
            for c in range(rg["min_col"], rg["max_col"] + 1):
                if (r, c) != (rg["min_row"], rg["min_col"]):
                    skip.add((r, c))
                if is_horizontal and c != rg["min_col"]:
                    hskip_cols_by_row.setdefault(r, set()).add(c)
    return skip, top_left, hskip_cols_by_row

def fast_get_merge_ranges(path):
    """
    Read merged ranges directly from the XLSX XML structure instead of
    loading the workbook normally with openpyxl.

    Returns:
        List of merge ranges on success.

        None if anything unexpected is encountered, allowing the caller
        to fall back to the standard openpyxl implementation.
    """
    import zipfile
    import re
    from xml.etree import ElementTree as ET

    try:
        with zipfile.ZipFile(path) as z:
            names = set(z.namelist())
            if "xl/workbook.xml" not in names:
                return None

            wb_xml = z.read("xl/workbook.xml").decode("utf-8")
            wb_root = ET.fromstring(wb_xml)
            ns = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
            ns_r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

            sheets = wb_root.findall("m:sheets/m:sheet", ns)
            if not sheets:
                return None

            # Determine the active worksheet. If activeTab is not available, use the first sheet.
            active_idx = 0
            book_views = wb_root.find("m:bookViews/m:workbookView", ns)
            if book_views is not None and book_views.get("activeTab") is not None:
                try:
                    active_idx = int(book_views.get("activeTab"))
                except ValueError:
                    active_idx = 0
            if active_idx >= len(sheets):
                active_idx = 0

            r_id = sheets[active_idx].get(f"{{{ns_r}}}id")
            if not r_id:
                return None

            rels_path = "xl/_rels/workbook.xml.rels"
            if rels_path not in names:
                return None
            rels_xml = z.read(rels_path).decode("utf-8")
            rels_root = ET.fromstring(rels_xml)
            target = None
            for rel in rels_root:
                if rel.get("Id") == r_id:
                    target = rel.get("Target")
                    break
            if not target:
                return None

            target = target.lstrip("/")
            if not target.startswith("xl/"):
                target = "xl/" + target
            if target not in names:
                return None

            sheet_xml = z.read(target).decode("utf-8")

        # mergeCells is usually located near the end of the worksheet XML, but the entire document is scanned for safety.
        refs = re.findall(r'<mergeCell[^>]*\sref="([A-Z]+\d+):([A-Z]+\d+)"', sheet_xml)

        def col_to_num(col_str):
            num = 0
            for ch in col_str:
                num = num * 26 + (ord(ch) - ord("A") + 1)
            return num

        ranges = []
        for start, end in refs:
            sm = re.match(r"([A-Z]+)(\d+)", start)
            em = re.match(r"([A-Z]+)(\d+)", end)
            if not sm or not em:
                return None
            ranges.append(
                {
                    "min_col": col_to_num(sm.group(1)),
                    "min_row": int(sm.group(2)),
                    "max_col": col_to_num(em.group(1)),
                    "max_row": int(em.group(2)),
                }
            )
        return ranges
    except Exception:
        return None


def render(rows):
    if not rows:
        return ""
    widths = [0] * max(len(r) for r in rows)
    for row in rows:
        for i, v in enumerate(row):
            widths[i] = max(widths[i], len(v))
    border = "+" + "+".join("-" * (w + 2) for w in widths) + "+"
    out = [border]
    for row in rows:
        parts = ["|"]
        for i, v in enumerate(row):
            parts.append(" ")
            parts.append(v.ljust(widths[i]))
            parts.append(" |")
        out.append("".join(parts))
        out.append(border)
    return "\n".join(out)


def open_xlsx(path):
    ensure_workbook(path)

    # Read merge ranges directly from XML. Fall back to openpyxl if XML parsing fails.
    ranges = fast_get_merge_ranges(path)
    used_fast_path = ranges is not None

    if used_fast_path:
        wb = load_workbook(path, data_only=True, keep_links=False)
        ws = wb.active

        max_row = ws.max_row or 1
        max_col = ws.max_column or 1
    else:
        wb = load_workbook(path, data_only=True, keep_links=False)
        ws = wb.active
        max_row = ws.max_row
        max_col = ws.max_column
        ranges = get_merge_ranges(ws)

    skip, top_left, hskip_cols_by_row = build_display_plan(ranges, max_row, max_col)

    display_rows = []

    for r, row in enumerate(ws.iter_rows(min_row=1, max_row=max_row), start=1):
        cols_to_drop = hskip_cols_by_row.get(r)
        display_row = []
        row_len = len(row)
        for c in range(1, max_col + 1):
            if cols_to_drop and c in cols_to_drop:
                continue
            if (r, c) in skip:
                display_row.append("")
            else:
                v = row[c - 1].value if c <= row_len else None
                display_row.append(
                    ""
                    if v is None
                    else str(v).replace("\r", " ").replace("\n", " \\n ")
                )
        display_rows.append(display_row)

    if used_fast_path:
        wb.close()

    if not display_rows:
        display_rows = [[""]]

    print(render(display_rows))


def parse_ascii(lines):
    rows = []
    for line in lines:
        line = line.rstrip()
        if not line.startswith("|"):
            continue
        cols = line.split("|")[1:-1]
        rows.append([c.strip() for c in cols])
    return rows


def compute_mergeinfo(path):
    # Calculate merge ranges and column mapping directly from the existing .xlsx file, using the same logic as open_xlsx. No caching to disk -> no temporary files generated.
    ranges = fast_get_merge_ranges(path)
    if ranges is None:
        wb = load_workbook(path, data_only=True, keep_links=False)
        ws = wb.active
        ranges = get_merge_ranges(ws)
        max_row, max_col = ws.max_row, ws.max_column
        wb.close()
    else:
        wb = load_workbook(path, read_only=True, data_only=True, keep_links=False)
        ws = wb.active
        max_row, max_col = ws.max_row or 1, ws.max_column or 1
        wb.close()

    skip, _, hskip_cols_by_row = build_display_plan(ranges, max_row, max_col)

    col_map_per_row = []
    for r in range(1, max_row + 1):
        cols_to_drop = hskip_cols_by_row.get(r, set())
        col_map = [c for c in range(1, max_col + 1) if c not in cols_to_drop]
        col_map_per_row.append(col_map)

    return {
        "max_row": max_row,
        "max_col": max_col,
        "ranges": ranges,
        "col_map_per_row": col_map_per_row,
    }


def save_xlsx(xlsx_file, txt_file):
    ensure_workbook(xlsx_file)
    with open(txt_file, encoding="utf-8") as f:
        lines = f.readlines()
    new_rows = parse_ascii(lines)

    info = compute_mergeinfo(xlsx_file)

    wb = load_workbook(xlsx_file)
    ws = wb.active

    for mc in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(mc))

    if len(info["col_map_per_row"]) >= len(new_rows):
        col_map_per_row = info["col_map_per_row"]
    else:
        # fallback: số dòng không khớp -> map 1:1 như cũ
        col_map_per_row = [list(range(1, len(row) + 1)) for row in new_rows]

    for r_idx, row in enumerate(new_rows):
        col_map = (
            col_map_per_row[r_idx]
            if r_idx < len(col_map_per_row)
            else list(range(1, len(row) + 1))
        )
        for i, value in enumerate(row):
            if i >= len(col_map):
                continue
            real_col = col_map[i]
            value = value.replace("\\n", "\n")
            cell = ws.cell(r_idx + 1, real_col)
            if str(cell.value or "") != value:
                cell.value = value

    # Áp lại các vùng merge gốc
    for rg in info["ranges"]:
        try:
            start = f"{get_column_letter(rg['min_col'])}{rg['min_row']}"
            end = f"{get_column_letter(rg['max_col'])}{rg['max_row']}"
            ws.merge_cells(f"{start}:{end}")
        except Exception:
            pass

    wb.save(xlsx_file)


def shift_merges_for_inserted_row(old_ranges, row_num):
    new_ranges = []
    for rg in old_ranges:
        min_row, max_row = rg["min_row"], rg["max_row"]
        if min_row >= row_num:
            min_row += 1
            max_row += 1
        elif max_row >= row_num:
            # vùng merge bao trùm vị trí chèn -> mở rộng thêm 1 dòng
            max_row += 1
        new_ranges.append(
            {
                "min_row": min_row,
                "max_row": max_row,
                "min_col": rg["min_col"],
                "max_col": rg["max_col"],
            }
        )
    return new_ranges


def shift_merges_for_inserted_col(old_ranges, col_num):
    new_ranges = []
    for rg in old_ranges:
        min_col, max_col = rg["min_col"], rg["max_col"]
        if min_col >= col_num:
            min_col += 1
            max_col += 1
        elif max_col >= col_num:
            max_col += 1
        new_ranges.append(
            {
                "min_row": rg["min_row"],
                "max_row": rg["max_row"],
                "min_col": min_col,
                "max_col": max_col,
            }
        )
    return new_ranges


def insert_row(path, row_num):
    wb = load_workbook(path)
    ws = wb.active
    old_ranges = get_merge_ranges(ws)
    for mc in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(mc))
    ws.insert_rows(row_num)
    max_cols = ws.max_column
    for c in range(1, max_cols + 1):
        ws.cell(row_num, c, "")
    new_ranges = shift_merges_for_inserted_row(old_ranges, row_num)
    for rg in new_ranges:
        start = f"{get_column_letter(rg['min_col'])}{rg['min_row']}"
        end = f"{get_column_letter(rg['max_col'])}{rg['max_row']}"
        try:
            ws.merge_cells(f"{start}:{end}")
        except Exception:
            pass
    wb.save(path)


def insert_col(path, col_num):
    wb = load_workbook(path)
    ws = wb.active
    old_ranges = get_merge_ranges(ws)
    for mc in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(mc))
    ws.insert_cols(col_num)
    max_rows = ws.max_row
    for r in range(1, max_rows + 1):
        ws.cell(r, col_num, "")
    new_ranges = shift_merges_for_inserted_col(old_ranges, col_num)
    for rg in new_ranges:
        start = f"{get_column_letter(rg['min_col'])}{rg['min_row']}"
        end = f"{get_column_letter(rg['max_col'])}{rg['max_row']}"
        try:
            ws.merge_cells(f"{start}:{end}")
        except Exception:
            pass
    wb.save(path)


if __name__ == "__main__":
    mode = sys.argv[1]
    if mode == "open":
        open_xlsx(sys.argv[2])
    elif mode == "save":
        save_xlsx(sys.argv[2], sys.argv[3])
    elif mode == "insert_row":
        insert_row(sys.argv[2], int(sys.argv[3]))
    elif mode == "insert_col":
        insert_col(sys.argv[2], int(sys.argv[3]))
