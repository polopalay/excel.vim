" ============================================================================
" GUARD & KHAI BÁO BIẾN GỐC
" ============================================================================
" Ngăn plugin load lại nhiều lần 
if exists('g:loaded_excelPlugin')
    finish
endif
let g:loaded_excelPlugin = 1
" Xác định đường dẫn gốc của plugin (lùi lên 1 cấp từ thư mục chứa file này)
" Ví dụ: nếu file này nằm ở plugin/excelPlugin.vim thì root_dir là thư mục cha
let s:plugin_root = fnamemodify(expand('<sfile>:p:h'), ':h')
let s:rs_root = s:plugin_root . '/rs'
" Xác định đường dẫn tới binary excel đã build (release)
if has('win32') || has('win64')
    let s:excel_bin = s:rs_root . '/target/release/excel.exe'
else
    let s:excel_bin = s:rs_root . '/target/release/excel'
endif
" ============================================================================
" INIT - AUTOCMD ĐĂNG KÝ MỞ FILE / FILETYPE / OVERRIDE ZIP
" ============================================================================
" ----------------------------------------------------------------------------
" Autocmd chính: khi mở file *.xlsx thì gọi hàm ExcelOpen() để xử lý
" thay vì để Vim mở file theo cách thông thường (binary/zip)
" ----------------------------------------------------------------------------
augroup ExcelViewer
    autocmd!
    autocmd BufReadCmd *.xlsx call ExcelOpen()
augroup END
" ----------------------------------------------------------------------------
" Đảm bảo ExcelOpen() được ưu tiên hơn cơ chế xử lý zip mặc định của Vim
" (vì file .xlsx thực chất là 1 file zip, nên plugin zip.vim built-in của Vim
" cũng đăng ký autocmd BufReadCmd cho *.xlsx và có thể tranh quyền xử lý)
" ----------------------------------------------------------------------------
function! s:EnsureExcelOverridesZip() abort
    " Xoá autocmd BufReadCmd của group 'zip' áp dụng riêng cho *.xlsx (nếu có)
    if exists('#zip#BufReadCmd#*.xlsx')
        autocmd! zip BufReadCmd *.xlsx
    endif
    " Xoá autocmd BufWriteCmd của group 'zip' áp dụng riêng cho *.xlsx (nếu có)
    if exists('#zip#BufWriteCmd#*.xlsx')
        autocmd! zip BufWriteCmd *.xlsx
    endif
endfunction
" Gọi hàm trên ở thời điểm VimEnter và SourcePost để đảm bảo zip.vim
" (thường autoload muộn) đã có cơ hội đăng ký autocmd trước khi ta xoá
augroup ExcelViewerInit
    autocmd!
    autocmd VimEnter,SourcePost * call s:EnsureExcelOverridesZip()
augroup END
" ----------------------------------------------------------------------------
" Tự động gán filetype=excel cho mọi file .xlsx, dù mở mới hay đọc lại
" (bổ sung phòng trường hợp filetype chưa được set qua ExcelOpen)
" ----------------------------------------------------------------------------
augroup ExcelFiletype
    autocmd!
    autocmd BufRead,BufNewFile *.xlsx setfiletype excel
augroup END
" ----------------------------------------------------------------------------
" Đảm bảo binary excel đã được build trước khi dùng, nếu chưa thì build
" ----------------------------------------------------------------------------
function! s:EnsureBuilt() abort
    if filereadable(s:excel_bin)
        return 1
    endif
    echo 'First build excel...'
	if executable('cargo') == 0
    	echoerr 'cargo not found in PATH'
    	return 0
	endif
    let l:cmd =
                \ 'cargo build --release --manifest-path '
                \ . shellescape(s:rs_root . '/Cargo.toml')
    let l:out = systemlist(l:cmd)
    if v:shell_error
        echoerr join(l:out, "\n")
        return 0
    endif
    return filereadable(s:excel_bin)
endfunction
" ============================================================================
" HELPER FUNCTIONS
" ============================================================================
" ----------------------------------------------------------------------------
" Tô màu cú pháp (syntax highlight) cho nội dung file Excel hiển thị dạng bảng
" ----------------------------------------------------------------------------
function! ExcelSetupHighlight()
    " Xoá hết các match cũ trước khi setup lại (tránh chồng nhiều lần)
    silent! call clearmatches()
    " Định nghĩa các highlight group riêng cho Excel
    highlight ExcelBorder guifg=#9ca0b0 ctermfg=245
    highlight ExcelDate   guifg=#40a02b ctermfg=70
    highlight ExcelLink   guifg=#0000ff ctermfg=12 gui=underline cterm=underline
    " ExcelNumber dùng chung style với group Number có sẵn của Vim
    highlight default link ExcelNumber Number
    " --- Viền bảng: +, |, dãy dấu - ---
    call matchadd('ExcelBorder', '[+]', 5)
    call matchadd('ExcelBorder', '[|]', 5)
    call matchadd('ExcelBorder', '-\+', 5)
    " --- URL: http/https/ftp ---
    call matchadd(
            \ 'ExcelLink',
            \ '\v%(https?|ftp)://[[:alnum:]/._~:?#[\]@!$&''()*+,;=%-]+',
            \ 30
            \ )
    " --- Số nguyên hoặc số thực ---
    call matchadd(
            \ 'ExcelNumber',
            \ '\v<\d+(\.\d+)?>',
            \ 15
            \ )
    " --- Ngày tháng: hỗ trợ nhiều định dạng phổ biến ---
    " dd/mm/yyyy hoặc dd/mm/yy (có thể kèm giờ:phút:giây)
    " mm/yyyy
    " yyyy-mm-dd (có thể kèm giờ:phút:giây)
    " dd-mm-yyyy
    call matchadd(
            \ 'ExcelDate',
            \ '\v<('
            \ . '\d{1,2}/\d{1,2}/\d{2,4}(\s+\d{1,2}:\d{2}(:\d{2})?)?'
            \ . '|'
            \ . '\d{1,2}/\d{4}'
            \ . '|'
            \ . '\d{4}-\d{1,2}-\d{1,2}(\s+\d{1,2}:\d{2}(:\d{2}(\.\d+)?)?)?'
            \ . '|'
            \ . '\d{1,2}-\d{1,2}-\d{4}'
            \ . ')>',
            \ 20
            \ )
    call matchadd(
            \ 'ExcelDate',
            \ '\v<\d{4}-\d{1,2}-\d{1,2}([ T]\d{1,2}:\d{2}(:\d{2})?(\.\d+)?)?>',
            \ 20
            \ )
    " --- Style thực từ file Excel (bold/italic/màu chữ/màu nền), parse từ
    " khối metadata @@STYLE@@ do binary excel xuất kèm sau bảng ASCII ---
    call s:ExcelApplyCellStyles()
endfunction
" ----------------------------------------------------------------------------
" s:ExcelHighlightGroupName(): sinh tên highlight group duy nhất, ổn định
" theo tổ hợp (bold, italic, font_hex, bg_hex), để không tạo trùng group
" cho cùng 1 kiểu style (Excel thường có rất nhiều cell dùng chung style).
" ----------------------------------------------------------------------------
function! s:ExcelHighlightGroupName(bold, italic, font_hex, bg_hex) abort
    let l:key = a:bold . a:italic . '_' . a:font_hex . '_' . a:bg_hex
    " Tên group Vim không được chứa ký tự đặc biệt như '#' hay '-'
    let l:key = substitute(l:key, '[^A-Za-z0-9_]', '_', 'g')
    return 'ExcelCellStyle_' . l:key
endfunction
" ----------------------------------------------------------------------------
" s:ExcelDefineHighlight(): định nghĩa (hoặc tái sử dụng nếu đã có) 1
" highlight group cho 1 tổ hợp style cụ thể.
" ----------------------------------------------------------------------------
function! s:ExcelDefineHighlight(group, bold, italic, font_hex, bg_hex) abort
    let l:attrs = []
    if a:bold
        call add(l:attrs, 'bold')
    endif
    if a:italic
        call add(l:attrs, 'italic')
    endif
    let l:gui_attr = empty(l:attrs) ? 'NONE' : join(l:attrs, ',')
    let l:cmd = 'highlight ' . a:group . ' gui=' . l:gui_attr . ' cterm=' . l:gui_attr
    if a:font_hex !=# '-'
        let l:cmd .= ' guifg=#' . a:font_hex
        let l:cmd .= ' ctermfg=' . s:ExcelHexToCterm(a:font_hex)
    endif
    if a:bg_hex !=# '-'
        let l:cmd .= ' guibg=#' . a:bg_hex
        let l:cmd .= ' ctermbg=' . s:ExcelHexToCterm(a:bg_hex)
    endif
    execute l:cmd
endfunction
" ----------------------------------------------------------------------------
" s:ExcelHexToCterm(): quy đổi gần đúng "RRGGBB" -> mã màu 256 (cterm) cho
" terminal không hỗ trợ guifg/guibg trực tiếp. Dùng công thức xterm-256
" chuẩn (6x6x6 color cube + grayscale ramp).
" ----------------------------------------------------------------------------
function! s:ExcelHexToCterm(hex) abort
    let l:r = str2nr(a:hex[0:1], 16)
    let l:g = str2nr(a:hex[2:3], 16)
    let l:b = str2nr(a:hex[4:5], 16)
    " Quy mỗi kênh 0-255 về 0-5 (6 mức của color cube xterm-256)
    let l:r6 = (l:r * 5 + 127) / 255
    let l:g6 = (l:g * 5 + 127) / 255
    let l:b6 = (l:b * 5 + 127) / 255
    return 16 + (36 * l:r6) + (6 * l:g6) + l:b6
endfunction
" Cache các highlight group đã định nghĩa trong session, tránh gọi lại
" :highlight nhiều lần không cần thiết cho cùng 1 tổ hợp style.
let s:excel_hl_defined = {}
" ----------------------------------------------------------------------------
" s:ExcelApplyCellStyles(): đọc b:xlsx_style_meta (đã parse từ khối
" @@STYLE@@...@@END@@ trong output của lệnh open/save), tạo highlight group
" tương ứng và gọi matchaddpos() để tô đúng vùng buffer.
"
" Lưu ý quan trọng về cột: binary excel tính col_start/col_end theo SỐ KÝ TỰ
" Unicode (char index, để khớp độ rộng cột ASCII table có dấu tiếng Việt),
" nhưng matchaddpos() của Vim lại cần BYTE index trong dòng. Nên phải
" convert qua byteidx() trước khi gọi, nếu không vùng tô sẽ bị lệch khi dòng
" có ký tự tiếng Việt nằm trước vị trí cần tô.
" ----------------------------------------------------------------------------
function! s:ExcelApplyCellStyles() abort
    if !exists('b:xlsx_style_meta') || empty(b:xlsx_style_meta)
        return
    endif
    for l:item in b:xlsx_style_meta
        let [l:line, l:char_start, l:char_end, l:bold, l:italic, l:font_hex, l:bg_hex] = l:item
        if l:line < 1 || l:line > line('$')
            continue
        endif
        let l:text = getline(l:line)
        " char_start/char_end là 1-based, inclusive (char index). byteidx()
        " nhận index 0-based -> char_start - 1 là ký tự đầu tiên cần tô.
        let l:byte_start = byteidx(l:text, l:char_start - 1)
        let l:byte_end = byteidx(l:text, l:char_end)
        if l:byte_start < 0 || l:byte_end < 0
            continue
        endif
        let l:byte_len = l:byte_end - l:byte_start
        if l:byte_len <= 0
            continue
        endif
        let l:group = s:ExcelHighlightGroupName(l:bold, l:italic, l:font_hex, l:bg_hex)
        if !get(s:excel_hl_defined, l:group, 0)
            call s:ExcelDefineHighlight(l:group, l:bold, l:italic, l:font_hex, l:bg_hex)
            let s:excel_hl_defined[l:group] = 1
        endif
        " matchaddpos nhận col theo byte (1-based) -> +1 vì byteidx() là 0-based
        call matchaddpos(l:group, [[l:line, l:byte_start + 1, l:byte_len]], 50)
    endfor
endfunction
" ----------------------------------------------------------------------------
" s:ExcelCellRefAt(line, vcol): tìm cell Excel (ví dụ "B2") chứa vị trí
" (line, vcol) trong buffer, dựa vào b:xlsx_cell_map đã parse từ @@CELLMAP@@.
"
" QUAN TRỌNG: `vcol` phải là VIRTUAL COLUMN (kết quả của virtcol(), KHÔNG
" phải col()/byte column), và so sánh trực tiếp với char_start/char_end từ
" Rust — vì bảng ASCII không chứa tab và không có ký tự full-width (CJK),
" nên virtual column trùng đúng với char index. Đây là điểm mấu chốt để xử
" lý đúng các dòng có độ dài byte khác nhau do dấu tiếng Việt: byte offset
" của "cùng 1 cột nhìn thấy" lệch nhau giữa các dòng, nhưng virtual column
" (và char index) luôn nhất quán. KHÔNG dùng byteidx()/col() để so sánh vị
" trí cell ở đây; byteidx() chỉ dùng khi gọi matchaddpos() (xem
" s:ExcelApplyCellStyles), vì đó là API duy nhất bắt buộc byte index.
" ----------------------------------------------------------------------------
function! s:ExcelCellRefAt(line, vcol) abort
    if !exists('b:xlsx_cell_map') || empty(b:xlsx_cell_map)
        return ''
    endif
    for l:item in b:xlsx_cell_map
        let [l:line_no, l:char_start, l:char_end, l:cell_ref] = l:item
        if l:line_no != a:line
            continue
        endif
        if a:vcol >= l:char_start && a:vcol <= l:char_end
            return l:cell_ref
        endif
    endfor
    return ''
endfunction
" ----------------------------------------------------------------------------
" s:ExcelCellAtCursor(): trả về địa chỉ cell Excel tại vị trí con trỏ hiện
" tại trong buffer, hoặc '' nếu con trỏ không nằm trong cell nào (dòng viền,
" ngoài bảng, hoặc file không phải buffer Excel).
" ----------------------------------------------------------------------------
function! s:ExcelCellAtCursor() abort
    return s:ExcelCellRefAt(line('.'), virtcol('.'))
endfunction
" ----------------------------------------------------------------------------
" s:ExcelCellsInVisualSelection(line1, line2): trả về danh sách (không trùng
" lặp) các địa chỉ cell Excel nằm trong vùng [line1,line2] x [vcol1,vcol2].
"
" line1/line2 là source of truth cho dòng (từ <line1>/<line2> của command).
" Cột thì cố lấy từ marks '<'/'>' (chỉ khi marks khớp đúng line1/line2 —
" trường hợp gọi từ Visual mode). Nếu marks không khớp (ví dụ user gõ
" :5,10ExcelBold), fallback line-wise (toàn bộ cell trong các dòng đó).
" ----------------------------------------------------------------------------
function! s:ExcelCellsInVisualSelection(line1, line2) abort
    let l:marks_match = line("'<") == a:line1 && line("'>") == a:line2
    let l:mode = visualmode()
    let l:is_block = l:marks_match && (l:mode ==# "\<C-v>")
    let l:is_linewise = (!l:marks_match) || l:mode ==# 'V'

    " Lấy column range nếu marks khớp, ngược lại bỏ qua check column.
    if l:marks_match
        let l:pos1 = getpos("'<")
        let l:pos2 = getpos("'>")
        let l:vcol1 = virtcol([a:line1, l:pos1[2]])
        let l:vcol2 = virtcol([a:line2, l:pos2[2]])
        " Nếu chạm '$' cuối dòng, kẹp lại theo cuối dòng cuối.
        let l:line2_vcol_max = virtcol([a:line2, '$']) - 1
        if l:vcol2 > l:line2_vcol_max
            let l:vcol2 = l:line2_vcol_max
        endif
        if l:is_block
            let l:block_col_min = min([l:vcol1, l:vcol2])
            let l:block_col_max = max([l:vcol1, l:vcol2])
        endif
    endif

    let l:refs = []
    let l:seen = {}
    if !exists('b:xlsx_cell_map')
        return l:refs
    endif
    for l:item in b:xlsx_cell_map
        let [l:line_no, l:char_start, l:char_end, l:cell_ref] = l:item
        if l:line_no < a:line1 || l:line_no > a:line2
            continue
        endif

        let l:in_range = 0
        if l:is_linewise
            " Line-wise: lấy mọi cell trong dòng, không xét cột.
            let l:in_range = 1
        elseif l:is_block
            let l:in_range = (l:char_end >= l:block_col_min && l:char_start <= l:block_col_max)
        elseif l:line_no > a:line1 && l:line_no < a:line2
            " Char-wise, dòng giữa: lấy toàn dòng.
            let l:in_range = 1
        elseif a:line1 == a:line2
            " Char-wise, cùng dòng: cell phải giao với [vcol1, vcol2].
            let l:in_range = (l:char_end >= l:vcol1 && l:char_start <= l:vcol2)
        elseif l:line_no == a:line1
            let l:in_range = (l:char_end >= l:vcol1)
        elseif l:line_no == a:line2
            let l:in_range = (l:char_start <= l:vcol2)
        endif

        if l:in_range && !has_key(l:seen, l:cell_ref)
            let l:seen[l:cell_ref] = 1
            call add(l:refs, l:cell_ref)
        endif
    endfor
    return l:refs
endfunction
" ----------------------------------------------------------------------------
" s:ExcelResolveTargetCells(line1, line2): trả về danh sách cell cần áp style.
"
" Cách phát hiện Visual mode vs Normal mode (KHÔNG dựa vào <range> value vì
" giá trị này có thể không đáng tin cậy giữa các phiên bản Vim):
"   - Nếu line1 != line2: chắc chắn user pass range -> Visual/range mode.
"   - Nếu line1 == line2 == current_line VÀ marks '< '> trùng đúng dòng đó
"     trên CÙNG vị trí cursor hiện tại: Normal mode (cursor).
"   - Trường hợp còn lại (line1 == line2 nhưng khớp marks '< '>):
"     Visual mode trên 1 dòng -> dùng marks để lấy cell theo cột.
"
" Báo lỗi và trả về [] nếu không xác định được cell nào (con trỏ ở dòng
" viền, ngoài bảng, hoặc Visual selection không chạm cell nào).
" ----------------------------------------------------------------------------
function! s:ExcelResolveTargetCells(line1, line2) abort
    " Logic phát hiện Visual mode:
    "
    "   1. Nếu line1 != line2 (range nhiều dòng) -> luôn là Visual/range.
    "   2. Nếu line1 == line2 (cùng 1 dòng): có thể là Normal mode tại con
    "      trỏ, HOẶC Visual char-wise trên 1 dòng. Phân biệt qua marks:
    "      nếu '<'/'>' khớp đúng line1 và cột virtcol('<')/virtcol('>') KHÁC
    "      nhau (selection có độ rộng > 1 ký tự) -> Visual. Còn lại Normal.
    "
    " Cách này tránh hoàn toàn việc phụ thuộc <range> value (vốn có thể
    " không nhất quán giữa các phiên bản Vim).
    let l:is_visual = 0
    if a:line1 != a:line2
        let l:is_visual = 1
    elseif line("'<") == a:line1 && line("'>") == a:line2
                \ && virtcol("'<") != virtcol("'>")
        let l:is_visual = 1
    endif

    if l:is_visual
        let l:refs = s:ExcelCellsInVisualSelection(a:line1, a:line2)
        if empty(l:refs)
            echoerr 'Selection does not cover any cell'
            return []
        endif
        return l:refs
    endif
    let l:ref = s:ExcelCellAtCursor()
    if empty(l:ref)
        echoerr 'Cursor is not inside any cell'
        return []
    endif
    return [l:ref]
endfunction
" ----------------------------------------------------------------------------
" s:ExcelParseStyleMeta(): tách khối @@STYLE@@...@@END@@,
" @@CELLMAP@@...@@CELLMAPEND@@ và @@FORMULA@@...@@FORMULAEND@@ ra khỏi
" output thô của binary excel.
" Trả về [bang_ascii, danh_sach_style, danh_sach_cellmap, dict_formula].
" - Mỗi phần tử style: [line, col_start, col_end, bold, italic, font_hex, bg_hex]
" - Mỗi phần tử cellmap: [line, col_start, col_end, cell_ref]
" - dict_formula: {cell_ref: formula_text} (formula_text có dấu "=" đầu)
" với line/col_start/col_end/bold/italic đã convert sang Number.
" ----------------------------------------------------------------------------
function! s:ExcelParseStyleMeta(output) abort
    let l:marker_idx = index(a:output, '@@STYLE@@')
    if l:marker_idx == -1
        return [a:output, [], [], {}]
    endif
    let l:table_lines = a:output[0 : l:marker_idx - 1]
    " Tìm '@@END@@' chỉ trong phần SAU marker (tự cắt list trước khi index()
    " để không phụ thuộc tham số {start} của index(), giữ tương thích rộng).
    let l:after_marker = a:output[l:marker_idx + 1 :]
    let l:end_idx_rel = index(l:after_marker, '@@END@@')
    if l:end_idx_rel == -1
        let l:style_lines = l:after_marker
        let l:after_style_end = []
    else
        let l:style_lines = l:after_marker[0 : l:end_idx_rel - 1]
        let l:after_style_end = l:after_marker[l:end_idx_rel + 1 :]
    endif

    let l:styles = []
    for l:raw in l:style_lines
        if empty(l:raw)
            continue
        endif
        let l:parts = split(l:raw, "\t")
        if len(l:parts) != 7
            continue
        endif
        call add(l:styles, [
                \ str2nr(l:parts[0]),
                \ str2nr(l:parts[1]),
                \ str2nr(l:parts[2]),
                \ str2nr(l:parts[3]),
                \ str2nr(l:parts[4]),
                \ l:parts[5],
                \ l:parts[6],
                \ ])
    endfor

    " --- Parse khối @@CELLMAP@@...@@CELLMAPEND@@ (nếu có) ---
    let l:cellmap = []
    let l:after_cellmap_end = []
    let l:cm_marker_idx = index(l:after_style_end, '@@CELLMAP@@')
    if l:cm_marker_idx != -1
        let l:after_cm_marker = l:after_style_end[l:cm_marker_idx + 1 :]
        let l:cm_end_idx_rel = index(l:after_cm_marker, '@@CELLMAPEND@@')
        let l:cellmap_lines = l:cm_end_idx_rel == -1
                    \ ? l:after_cm_marker
                    \ : l:after_cm_marker[0 : l:cm_end_idx_rel - 1]
        let l:after_cellmap_end = l:cm_end_idx_rel == -1
                    \ ? []
                    \ : l:after_cm_marker[l:cm_end_idx_rel + 1 :]
        for l:raw in l:cellmap_lines
            if empty(l:raw)
                continue
            endif
            let l:parts = split(l:raw, "\t")
            if len(l:parts) != 4
                continue
            endif
            call add(l:cellmap, [
                    \ str2nr(l:parts[0]),
                    \ str2nr(l:parts[1]),
                    \ str2nr(l:parts[2]),
                    \ l:parts[3],
                    \ ])
        endfor
    endif

    " --- Parse khối @@FORMULA@@...@@FORMULAEND@@ (nếu có) ---
    " Dict cell_ref -> formula text (đã có dấu "=" đầu, vd "=SUM(A1:A4)"),
    " dùng để hiển thị formula khi con trỏ đứng trên 1 cell (xem
    " s:ExcelUpdateStatusCell / ExcelStatusLine).
    let l:formulas = {}
    let l:fm_marker_idx = index(l:after_cellmap_end, '@@FORMULA@@')
    if l:fm_marker_idx != -1
        let l:after_fm_marker = l:after_cellmap_end[l:fm_marker_idx + 1 :]
        let l:fm_end_idx_rel = index(l:after_fm_marker, '@@FORMULAEND@@')
        let l:formula_lines = l:fm_end_idx_rel == -1
                    \ ? l:after_fm_marker
                    \ : l:after_fm_marker[0 : l:fm_end_idx_rel - 1]
        for l:raw in l:formula_lines
            if empty(l:raw)
                continue
            endif
            " Chỉ tách ở tab ĐẦU TIÊN (formula text về lý thuyết không nên
            " chứa tab, nhưng phòng hờ vẫn an toàn hơn split() thường).
            let l:tab_idx = stridx(l:raw, "\t")
            if l:tab_idx == -1
                continue
            endif
            let l:formulas[l:raw[0 : l:tab_idx - 1]] = l:raw[l:tab_idx + 1 :]
        endfor
    endif

    " Loại bỏ dòng trống cuối bảng ASCII (giữa bảng và marker @@STYLE@@,
    " println! của Rust luôn thêm 1 dòng trống sau bảng).
    while !empty(l:table_lines) && l:table_lines[-1] ==# ''
        call remove(l:table_lines, -1)
    endwhile
    return [l:table_lines, l:styles, l:cellmap, l:formulas]
endfunction
" ----------------------------------------------------------------------------
" ExcelCmd(): hàm dùng chung để gọi binary excel với 1 mode
" (open/save/sheets/...) cùng các tham số bổ sung, trả về output dạng list
" các dòng (mỗi dòng 1 string), giống cách systemlist() hoạt động.
" ----------------------------------------------------------------------------
function! s:ExcelCmd(mode, ...) abort
    let l:cmd =
                \ shellescape(s:excel_bin)
                \ . ' '
                \ . a:mode
                \ . ' '
                \ . shellescape(b:xlsx_file)
    for arg in a:000
        let l:cmd .= ' ' . shellescape(arg)
    endfor
    return systemlist(l:cmd)
endfunction
" ----------------------------------------------------------------------------
" s:ExcelLoadIntoBuffer(): gọi binary excel với mode 'open' (hoặc bất kỳ mode
" nào trả ra bảng ASCII + metadata style), tách metadata, nạp bảng vào buffer
" và lưu lại metadata vào b:xlsx_style_meta để ExcelSetupHighlight() dùng.
" Dùng chung cho ExcelOpen/ExcelSave/ExcelSheetOpen để tránh lặp code.
" ----------------------------------------------------------------------------
function! s:ExcelLoadIntoBuffer(raw_output) abort
    let [l:table_lines, l:styles, l:cellmap, l:formulas] = s:ExcelParseStyleMeta(a:raw_output)
    let b:xlsx_style_meta = l:styles
    let b:xlsx_cell_map = l:cellmap
    let b:xlsx_formula_map = l:formulas
    setlocal modifiable
    silent %delete _
    call setline(1, l:table_lines)
    call ExcelSetupHighlight()
    call s:ExcelUpdateStatusCell()
endfunction
" ----------------------------------------------------------------------------
" s:ExcelReloadBuffer(sheet): gọi binary excel với mode 'open' cho sheet
" `sheet`, TỰ ĐỘNG kèm flag hiện công thức nếu b:xlsx_show_formulas đang
" bật (xem ExcelShowFormula()), rồi nạp kết quả vào buffer. Dùng CHUNG cho
" MỌI nơi cần "đã ghi file xong, giờ load lại để hiển thị bản chuẩn"
" (ExcelOpen, ExcelSave, ExcelSheetOpen, sau khi setbg/bold/merge/
" unmerge/applyformula...) — gom 1 chỗ để flag show-formula luôn được áp
" dụng nhất quán ở mọi điểm reload, không phải sửa rải rác từng hàm.
" Trả về 1 nếu thành công, 0 nếu lỗi (đã echoerr nội dung lỗi từ binary).
" ----------------------------------------------------------------------------
function! s:ExcelReloadBuffer(sheet) abort
    let l:flag = get(b:, 'xlsx_show_formulas', 0) ? '1' : '0'
    let l:output = s:ExcelCmd('open', a:sheet, l:flag)
    if v:shell_error
        echoerr join(l:output, "\n")
        return 0
    endif
    call s:ExcelLoadIntoBuffer(l:output)
    return 1
endfunction
" ----------------------------------------------------------------------------
" s:ExcelUpdateStatusCell(): cập nhật biến buffer-local `b:xlsx_status_cell`
" với địa chỉ ô Excel tại con trỏ hiện tại (vd "B3" hoặc "" nếu ngoài bảng),
" và `b:xlsx_status_formula` với formula của ô đó (vd "=SUM(A1:A4)", hoặc
" "" nếu ô không có công thức) — đây chính là cơ chế "hover" hiển thị
" formula: vì terminal Vim không có khái niệm hover bằng chuột, ta dùng lại
" đúng cơ chế cursor đã có sẵn của statusline (xem ExcelStatusLine()).
" Hàm này được gọi qua autocmd CursorMoved/CursorMovedI để statusline luôn
" hiển thị đúng vị trí. Dùng cache thay vì gọi trực tiếp trong statusline
" để tránh overhead — statusline có thể bị Vim redraw rất thường xuyên.
" ----------------------------------------------------------------------------
function! s:ExcelUpdateStatusCell() abort
    let b:xlsx_status_cell = s:ExcelCellAtCursor()
    let b:xlsx_status_formula = get(get(b:, 'xlsx_formula_map', {}), b:xlsx_status_cell, '')
endfunction
" ----------------------------------------------------------------------------
" ExcelStatusLine(): hàm gọi từ 'statusline' của Vim để hiển thị địa chỉ ô
" Excel tại con trỏ, kèm formula của ô đó nếu có. Format:
" "Sheet1 / B3 =SUM(A1:A4) — file.xlsx" (phần formula chỉ xuất hiện khi ô
" hiện tại có công thức). Trả về string đã thoát các ký tự đặc biệt để
" dùng trực tiếp trong %{...} của statusline.
" ----------------------------------------------------------------------------
function! ExcelStatusLine() abort
    if !exists('b:xlsx_file')
        return ''
    endif
    let l:sheet = exists('b:xlsx_sheet') ? b:xlsx_sheet : '?'
    let l:cell = get(b:, 'xlsx_status_cell', '')
    let l:cell_part = empty(l:cell) ? '(outside table)' : l:cell
    let l:formula = get(b:, 'xlsx_status_formula', '')
    let l:formula_part = empty(l:formula) ? '' : ('  ' . l:formula)
    return l:sheet . ' / ' . l:cell_part . l:formula_part
endfunction
" ----------------------------------------------------------------------------
" ExcelGoto(ref): nhảy con trỏ đến ô có địa chỉ Excel `ref` (vd "B3"). Tra
" trong b:xlsx_cell_map (tìm entry đầu tiên có cell_ref khớp) để biết
" line/char_start trong buffer, rồi cursor() đến đó. Báo lỗi nếu không tìm
" thấy (ô ngoài phạm vi hiển thị hiện tại).
" ----------------------------------------------------------------------------
function! ExcelGoto(ref) abort
    if !exists('b:xlsx_cell_map')
        echoerr 'This buffer is not an Excel file'
        return
    endif
    let l:target = toupper(trim(a:ref))
    for l:item in b:xlsx_cell_map
        let [l:line_no, l:char_start, l:char_end, l:cell_ref] = l:item
        if l:cell_ref ==# l:target
            " Convert char-index -> byte-index để cursor() đứng đúng vị trí
            " trên dòng có ký tự nhiều byte (tiếng Việt).
            let l:text = getline(l:line_no)
            let l:byte_col = byteidx(l:text, l:char_start - 1) + 1
            call cursor(l:line_no, l:byte_col)
            return
        endif
    endfor
    echoerr 'Cell not found: ' . a:ref
endfunction
" ----------------------------------------------------------------------------
" ExcelGotoComplete(): Tab-completion cho :ExcelGoto, liệt kê toàn bộ cell
" ref khả dụng trong cellmap hiện tại.
" ----------------------------------------------------------------------------
function! ExcelGotoComplete(A, L, P) abort
    if !exists('b:xlsx_cell_map')
        return []
    endif
    let l:refs = []
    let l:seen = {}
    for l:item in b:xlsx_cell_map
        let l:ref = l:item[3]
        if !has_key(l:seen, l:ref)
            let l:seen[l:ref] = 1
            call add(l:refs, l:ref)
        endif
    endfor
    return filter(l:refs, 'v:val =~? "^" . a:A')
endfunction
" ----------------------------------------------------------------------------
" ExcelSheetComplete(): hàm completion dùng cho các command có
" -complete=customlist (gợi ý tên sheet khi user gõ :ExcelSheetOpen <Tab>)
" ----------------------------------------------------------------------------
function! ExcelSheetComplete(A, L, P) abort
    if !exists('b:xlsx_file')
        return []
    endif
    return filter(
                \ s:ExcelCmd('sheets'),
                \ 'v:val =~? "^" . a:A'
                \ )
endfunction
" ============================================================================
" CORE FUNCTIONS - MỞ / LƯU / QUẢN LÝ SHEET
" ============================================================================
" ----------------------------------------------------------------------------
" ExcelBuild(): build binary excel nếu chưa có
" ----------------------------------------------------------------------------
function! ExcelBuild() abort
    if executable('cargo') == 0
        echoerr 'cargo not found in PATH'
        return
    endif
    echo 'Building excel_rs...'
    let l:cmd =
                \ 'cargo build --release --manifest-path '
                \ . shellescape(s:rs_root . '/Cargo.toml')
    let l:out = systemlist(l:cmd)
    if v:shell_error
        echoerr join(l:out, "\n")
        return
    endif
    echo 'Build success: ' . s:excel_bin
endfunction
" ----------------------------------------------------------------------------
" ExcelOpen(): đọc file .xlsx qua binary excel, hiển thị nội dung dạng bảng text trong buffer hiện tại
" ----------------------------------------------------------------------------
function! ExcelOpen()
    " Tránh gọi đệ quy (vì hàm này có thể tự gọi lại buffer khi reload)
    if exists('b:xlsx_buffer')
        return
    endif
	if !s:EnsureBuilt()
    	echoerr 'Build failed'
    	return
	endif
    let b:xlsx_buffer = 1
    " Lấy đường dẫn file vừa được match bởi autocmd
    let l:file = expand('<amatch>')
    " Nếu không lấy được (trường hợp gọi thủ công, không qua autocmd)
    " thì lấy từ tên file của buffer hiện tại
    if empty(l:file)
        let l:file = expand('%:p')
    endif
    " Lưu đường dẫn tuyệt đối của file .xlsx vào biến buffer-local
    let b:xlsx_file = fnamemodify(l:file, ':p')
    if !exists('b:xlsx_sheet')
        let b:xlsx_sheet = ''
    endif
    " Kiểm tra binary excel có tồn tại không
    if !executable(s:excel_bin) && !filereadable(s:excel_bin)
        echoerr 'excel binary not found: ' . s:excel_bin
        unlet! b:xlsx_buffer
        return
    endif
    " Gọi binary excel với lệnh 'open' để chuyển .xlsx -> text dạng bảng
    if !s:ExcelReloadBuffer(b:xlsx_sheet)
        unlet! b:xlsx_buffer
        return
    endif
    " Thiết lập các option cho buffer:
    " - buftype=acwrite: buffer không gắn trực tiếp với file thật,
    "   việc ghi (write) sẽ được xử lý thủ công qua autocmd BufWriteCmd
    " - bufhidden=hide: ẩn buffer khi đóng tab/window thay vì xoá hẳn
    " - noswapfile: không tạo swap file (vì đây là buffer "ảo")
    " - filetype=excel: để áp dụng highlight/ftplugin riêng cho Excel
    setlocal buftype=acwrite
    setlocal bufhidden=hide
    setlocal noswapfile
    setlocal filetype=excel
    " Statusline tùy chỉnh: hiển thị tên sheet + địa chỉ ô tại con trỏ (vd
    " "Sheet1 / B3"). Thông tin này KHÔNG nằm trong buffer text -> không bị
    " yank/visual copy theo, an toàn cho thao tác sao chép nội dung ô.
    setlocal statusline=%{ExcelStatusLine()}\ —\ %f%=%l,%c\ \ %P
    " Đăng ký autocmd lưu file: khi user gõ :w, gọi ExcelSave() thay vì
    " để Vim ghi buffer ra file theo cách thông thường
    augroup ExcelBuffer
        autocmd! * <buffer>
        autocmd BufWriteCmd <buffer> call ExcelSave()
        " Cập nhật cache địa chỉ ô khi con trỏ di chuyển — statusline sẽ
        " lấy giá trị này (qua ExcelStatusLine) thay vì tính lại mỗi lần
        " redraw, tránh overhead nếu cellmap lớn.
        autocmd CursorMoved,CursorMovedI <buffer> call s:ExcelUpdateStatusCell()
    augroup END
    " Đánh dấu buffer là "chưa chỉnh sửa" (vừa mở xong, chưa có thay đổi gì)
    set nomodified
endfunction
" ----------------------------------------------------------------------------
" ExcelSave(): lưu nội dung buffer (dạng bảng text) trở lại thành file .xlsx
" thông qua binary excel, sau đó load lại để hiển thị bản đã format chuẩn
" ----------------------------------------------------------------------------
function! ExcelSave()
    " Tạo file tạm để chứa nội dung hiện tại của buffer
    let l:tmp = tempname()
    call writefile(
                \ getline(1, '$'),
                \ l:tmp
                \ )
    " Gọi binary excel với lệnh 'save': đọc file tạm (dạng bảng text)
    " và ghi đè vào file .xlsx gốc
    let l:result = s:ExcelCmd(
            \ 'save',
            \ l:tmp,
            \ b:xlsx_sheet
            \ )
    " Xoá file tạm ngay sau khi dùng xong, không cần giữ lại
    call delete(l:tmp)
    " Nếu lưu lỗi thì báo lỗi và dừng (không reload lại buffer)
    if v:shell_error
        echoerr join(l:result, "\n")
        return
    endif
    " Sau khi lưu thành công, đọc lại file .xlsx để cập nhật buffer
    " với nội dung đã được excel format/chuẩn hoá lại (căn cột, v.v.)
    if !s:ExcelReloadBuffer(b:xlsx_sheet)
        return
    endif
    " Đánh dấu lại buffer là "đã lưu, không còn thay đổi"
    set nomodified
    echo 'Excel saved & reformatted'
endfunction
" ----------------------------------------------------------------------------
" ExcelSheets(): liệt kê tên tất cả các sheet trong file .xlsx hiện tại
" ----------------------------------------------------------------------------
function! ExcelSheets()
    echo join(
                \ s:ExcelCmd('sheets'),
                \ "\n"
                \ )
endfunction
" ----------------------------------------------------------------------------
" ExcelSheetOpen(): chuyển buffer hiện tại sang hiển thị 1 sheet khác
" ----------------------------------------------------------------------------
function! ExcelSheetOpen(sheet)
    if &modified
        set nomodified
    endif
    let b:xlsx_sheet = a:sheet
    if !s:ExcelReloadBuffer(a:sheet)
        return
    endif
    set nomodified
endfunction
" ----------------------------------------------------------------------------
" ExcelSheetAdd(): tạo 1 sheet mới và chuyển buffer sang hiển thị sheet đó
" ----------------------------------------------------------------------------
function! ExcelSheetAdd(sheet)
    call s:ExcelCmd(
                \ 'addsheet',
                \ a:sheet
                \ )
    let b:xlsx_sheet = a:sheet
    call ExcelSheetOpen(a:sheet)
    echo 'Sheet created: ' . a:sheet
endfunction
" ----------------------------------------------------------------------------
" ExcelSheetRename(): hiển thị menu chọn sheet rồi đổi tên sheet đó
" ----------------------------------------------------------------------------
function! ExcelSheetRename() abort
    let l:sheets = s:ExcelCmd('sheets')
    if empty(l:sheets)
        echoerr 'No sheets found'
        return
    endif
    let l:menu = ['Select sheet to rename:']
    for i in range(len(l:sheets))
        call add(l:menu, printf('%d. %s', i + 1, l:sheets[i]))
    endfor
    let l:old = inputlist(l:menu)
    if l:old < 1 || l:old > len(l:sheets)
        return
    endif
    let l:old_name = l:sheets[l:old - 1]
    let l:new_name = input(
                \ 'New sheet name: ',
                \ l:old_name
                \ )
    if empty(l:new_name)
        return
    endif
    let l:result = s:ExcelCmd(
                \ 'rensheet',
                \ l:old_name,
                \ l:new_name
                \ )
    if v:shell_error
        echoerr join(l:result, "\n")
        return
    endif
    if exists('b:xlsx_sheet')
                \ && b:xlsx_sheet ==# l:old_name
        let b:xlsx_sheet = l:new_name
    endif
    echo 'Sheet renamed: '
                \ . l:old_name
                \ . ' -> '
                \ . l:new_name
endfunction
" ----------------------------------------------------------------------------
" ExcelSheetDelete(): xoá 1 sheet sau khi xác nhận, rồi chuyển sang sheet
" đầu tiên còn lại (nếu có)
" ----------------------------------------------------------------------------
function! ExcelSheetDelete(sheet)
    if confirm(
                \ 'Delete sheet "' . a:sheet . '" ?',
                \ "&Yes\n&No",
                \ 2
                \ ) != 1
        return
    endif
    call s:ExcelCmd(
                \ 'delsheet',
                \ a:sheet
                \ )
    let l:sheets = s:ExcelCmd('sheets')
    if !empty(l:sheets)
        call ExcelSheetOpen(l:sheets[0])
    endif
    echo 'Sheet deleted: ' . a:sheet
endfunction
" ----------------------------------------------------------------------------
" Bộ màu chuẩn (tiếng Anh, khớp với Rust) dùng cho completion của
" :ExcelSetBg/:ExcelSetFg. "none" để xoá màu (trả về mặc định).
" ----------------------------------------------------------------------------
let s:excel_color_names = ['red', 'green', 'blue', 'yellow', 'orange', 'purple', 'gray', 'white', 'black', 'none']
" ----------------------------------------------------------------------------
" s:ExcelRunStyleCmd(): gọi binary excel với 1 mode style
" (setbg/setfg/togglebold/toggleitalic) cho TẤT CẢ cell trong `cell_refs` —
" join thành 1 chuỗi comma-separated rồi gọi binary 1 LẦN duy nhất (binary
" hỗ trợ "B2,C3,D5" hoặc "B2:D5" trong cùng 1 tham số), rồi reload buffer.
"
" Việc batch trong 1 lần gọi giúp:
"   - Nhanh hơn nhiều khi áp style cho 1 vùng lớn (N cell -> 1 process spawn
"     thay vì N spawn).
"   - Atomic: hoặc tất cả cell được áp style, hoặc không cell nào (nếu lỗi
"     ở giữa thì file không bị nửa vời).
" ----------------------------------------------------------------------------
function! s:ExcelRunStyleCmd(mode, cell_refs, extra_arg) abort
    if empty(a:cell_refs)
        return 0
    endif
    let l:joined = join(a:cell_refs, ',')
    let l:result = a:extra_arg ==# ''
                \ ? s:ExcelCmd(a:mode, l:joined, b:xlsx_sheet)
                \ : s:ExcelCmd(a:mode, l:joined, a:extra_arg, b:xlsx_sheet)
    if v:shell_error
        echoerr join(l:result, "\n")
        return 0
    endif
    if !s:ExcelReloadBuffer(b:xlsx_sheet)
        return 0
    endif
    set nomodified
    return 1
endfunction
" ----------------------------------------------------------------------------
" ExcelSetBg([color]): đổi màu nền cho cell tại con trỏ (Normal mode) hoặc
" toàn bộ cell trong vùng chọn (Visual mode, gọi qua :'<,'>ExcelSetBg).
" `color` là tên màu chuẩn (red/green/blue/yellow/orange/purple/gray/white/
" black) hoặc mã hex "#RRGGBB". Dùng "none" để xoá màu nền.
" ----------------------------------------------------------------------------
function! ExcelSetBg(color, line1, line2) abort
    let l:refs = s:ExcelResolveTargetCells(a:line1, a:line2)
    if empty(l:refs)
        return
    endif
    if s:ExcelRunStyleCmd('setbg', l:refs, a:color)
        echo 'Background color changed (' . len(l:refs) . ' cells): ' . join(l:refs, ', ') . ' -> ' . a:color
    endif
endfunction
" ----------------------------------------------------------------------------
" ExcelSetFg([color]): đổi màu chữ, tương tự ExcelSetBg.
" ----------------------------------------------------------------------------
function! ExcelSetFg(color, line1, line2) abort
    let l:refs = s:ExcelResolveTargetCells(a:line1, a:line2)
    if empty(l:refs)
        return
    endif
    if s:ExcelRunStyleCmd('setfg', l:refs, a:color)
        echo 'Font color changed (' . len(l:refs) . ' cells): ' . join(l:refs, ', ') . ' -> ' . a:color
    endif
endfunction
" ----------------------------------------------------------------------------
" ExcelBold([visual]): đảo trạng thái in đậm cho cell tại con trỏ hoặc toàn
" bộ vùng chọn. Mỗi cell được đảo ĐỘC LẬP (cell đang đậm sẽ tắt, cell đang
" không đậm sẽ bật) — giống đúng hành vi nút Bold trong Excel khi áp cho
" 1 vùng có cell trộn lẫn đậm/không đậm.
" ----------------------------------------------------------------------------
function! ExcelBold(line1, line2) abort
    let l:refs = s:ExcelResolveTargetCells(a:line1, a:line2)
    if empty(l:refs)
        return
    endif
    if s:ExcelRunStyleCmd('togglebold', l:refs, '')
        echo 'Bold toggled (' . len(l:refs) . ' cells): ' . join(l:refs, ', ')
    endif
endfunction
" ----------------------------------------------------------------------------
" ExcelItalic([visual]): đảo trạng thái in nghiêng, tương tự ExcelBold.
" ----------------------------------------------------------------------------
function! ExcelItalic(line1, line2) abort
    let l:refs = s:ExcelResolveTargetCells(a:line1, a:line2)
    if empty(l:refs)
        return
    endif
    if s:ExcelRunStyleCmd('toggleitalic', l:refs, '')
        echo 'Italic toggled (' . len(l:refs) . ' cells): ' . join(l:refs, ', ')
    endif
endfunction
" ----------------------------------------------------------------------------
" s:ExcelBoundingBoxRef(cell_refs): từ list cell ref ["B2","C3","A1"], tính
" ra "min_ref:max_ref" (vd "A1:C3") — bounding box bao trùm tất cả cell.
" Dùng để truyền 1 range duy nhất cho lệnh merge thay vì list cell rời.
" ----------------------------------------------------------------------------
function! s:ExcelBoundingBoxRef(cell_refs) abort
    if empty(a:cell_refs)
        return ''
    endif
    let l:min_row = 0
    let l:max_row = 0
    let l:min_col_letters = ''
    let l:max_col_letters = ''
    let l:min_col_n = 0
    let l:max_col_n = 0
    let l:first = 1
    for l:ref in a:cell_refs
        " Tách "B12" -> ("B", 12). Cột là phần chữ đầu, dòng là phần số sau.
        let l:m = matchlist(l:ref, '^\([A-Z]\+\)\(\d\+\)$')
        if empty(l:m)
            continue
        endif
        let l:letters = l:m[1]
        let l:row = str2nr(l:m[2])
        " Quy đổi cột chữ -> số: A=1, B=2, ..., Z=26, AA=27...
        let l:col_n = 0
        for l:ch in split(l:letters, '\zs')
            let l:col_n = l:col_n * 26 + (char2nr(l:ch) - char2nr('A') + 1)
        endfor
        if l:first
            let l:min_row = l:row
            let l:max_row = l:row
            let l:min_col_n = l:col_n
            let l:max_col_n = l:col_n
            let l:min_col_letters = l:letters
            let l:max_col_letters = l:letters
            let l:first = 0
        else
            if l:row < l:min_row | let l:min_row = l:row | endif
            if l:row > l:max_row | let l:max_row = l:row | endif
            if l:col_n < l:min_col_n
                let l:min_col_n = l:col_n
                let l:min_col_letters = l:letters
            endif
            if l:col_n > l:max_col_n
                let l:max_col_n = l:col_n
                let l:max_col_letters = l:letters
            endif
        endif
    endfor
    if l:first
        return ''
    endif
    return l:min_col_letters . l:min_row . ':' . l:max_col_letters . l:max_row
endfunction
" ----------------------------------------------------------------------------
" ExcelMerge(line1, line2): gộp các ô đã chọn (Visual mode) thành 1 vùng
" merge. Tính bounding box của Visual selection (min_row..max_row x
" min_col..max_col) rồi gọi binary mode 'merge'.
"
" Nếu gọi từ Normal mode (line1==line2 và không có visual selection thực
" sự), báo lỗi — vì merge 1 cell không có ý nghĩa.
" ----------------------------------------------------------------------------
function! ExcelMerge(line1, line2) abort
    let l:refs = s:ExcelResolveTargetCells(a:line1, a:line2)
    if empty(l:refs)
        return
    endif
    if len(l:refs) < 2
        echoerr 'Need to select at least 2 cells to merge (use Visual mode v/V/Ctrl-V)'
        return
    endif
    let l:bbox = s:ExcelBoundingBoxRef(l:refs)
    if empty(l:bbox)
        echoerr 'Could not determine merge range'
        return
    endif
    let l:result = s:ExcelCmd('merge', l:bbox, b:xlsx_sheet)
    if v:shell_error
        echoerr join(l:result, "\n")
        return
    endif
    if !s:ExcelReloadBuffer(b:xlsx_sheet)
        return
    endif
    set nomodified
    echo 'Merged range: ' . l:bbox
endfunction
" ----------------------------------------------------------------------------
" ExcelUnmerge(line1, line2): bỏ gộp 1 hoặc nhiều vùng merge.
" - Normal mode tại 1 ô trong vùng đã gộp: bỏ đúng vùng đó.
" - Visual mode: bỏ MỌI vùng merge giao với selection.
" ----------------------------------------------------------------------------
function! ExcelUnmerge(line1, line2) abort
    let l:refs = s:ExcelResolveTargetCells(a:line1, a:line2)
    if empty(l:refs)
        return
    endif
    " Với unmerge, ta truyền bounding box (nếu nhiều cell) hoặc cell đơn —
    " binary tự tìm và bỏ mọi merge giao với input đó.
    let l:input = len(l:refs) == 1 ? l:refs[0] : s:ExcelBoundingBoxRef(l:refs)
    let l:result = s:ExcelCmd('unmerge', l:input, b:xlsx_sheet)
    if v:shell_error
        echoerr join(l:result, "\n")
        return
    endif
    if !s:ExcelReloadBuffer(b:xlsx_sheet)
        return
    endif
    set nomodified
    echo 'Unmerged: ' . l:input
endfunction
" ----------------------------------------------------------------------------
" ExcelShowFormula(): bật/tắt mode hiển thị CÔNG THỨC ("=SUM(A1:A4)") ngay
" trong bảng thay vì giá trị đã tính — giống Ctrl+` trong Excel. Gọi lại
" lần 2 để quay về hiển thị giá trị bình thường. b:xlsx_show_formulas
" được lưu trên buffer nên giữ nguyên trạng thái qua mọi lần :w / đổi màu
" / merge... (mọi điểm reload đều đi qua s:ExcelReloadBuffer, tự đọc lại
" flag này mỗi lần).
"
" Khi mode đang BẬT, các cell có công thức hiển thị nguyên văn "=...".
" Sửa trực tiếp 1 trong các cell đó rồi :w hoạt động bình thường — vì nội
" dung hiển thị vẫn bắt đầu bằng "=", save_logic nhận diện y như khi gõ
" formula mới, không cần xử lý gì khác ở phía Rust.
" ----------------------------------------------------------------------------
function! ExcelShowFormula() abort
    if !exists('b:xlsx_file')
        echoerr 'This buffer is not an Excel file'
        return
    endif
    let b:xlsx_show_formulas = !get(b:, 'xlsx_show_formulas', 0)
    if !s:ExcelReloadBuffer(b:xlsx_sheet)
        " Lỗi -> rollback flag để lần gọi sau không bị kẹt ở trạng thái sai
        let b:xlsx_show_formulas = !b:xlsx_show_formulas
        return
    endif
    set nomodified
    echo b:xlsx_show_formulas ? 'Showing formulas' : 'Showing values'
endfunction
" ----------------------------------------------------------------------------
" ExcelApplyFormula(line1, line2): "Apply Formula" — tương đương kéo fill
" handle trong Excel. BẮT BUỘC gọi từ Visual mode (v/V/Ctrl-V), chọn 1
" vùng mà Ô ĐẦU TIÊN (trên cùng / bên trái nhất trong vùng) đang có công
" thức — ví dụ đang đứng ở B2 = "=B1+1", Visual chọn B2:B10 rồi gọi lệnh
" này: B3 sẽ tự thành "=B2+1", B4 thành "=B3+1", ... (mỗi ô dịch tham
" chiếu theo đúng độ lệch (dòng, cột) so với B2). Visual chọn ngang 1 dòng
" thì dịch theo cột tương tự. Visual chọn cả khối 2D cũng hoạt động vì độ
" lệch được tính riêng cho từng ô đích.
"
" Báo lỗi nếu: không ở Visual mode (chỉ có 1 ô), hoặc ô đầu tiên trong
" vùng chọn không có công thức sẵn.
" ----------------------------------------------------------------------------
function! ExcelApplyFormula(line1, line2) abort
    let l:refs = s:ExcelResolveTargetCells(a:line1, a:line2)
    if empty(l:refs)
        return
    endif
    if len(l:refs) < 2
        echoerr 'Cần Visual chọn ít nhất 2 ô: ô đầu (có công thức mẫu) + các ô áp dụng theo'
        return
    endif
    let l:result = s:ExcelCmd('applyformula', join(l:refs, ','), b:xlsx_sheet)
    if v:shell_error
        echoerr join(l:result, "\n")
        return
    endif
    if !s:ExcelReloadBuffer(b:xlsx_sheet)
        return
    endif
    set nomodified
    echo 'Applied formula from ' . l:refs[0] . ' to ' . (len(l:refs) - 1) . ' cell(s)'
endfunction
" ----------------------------------------------------------------------------
" ExcelColorComplete(): gợi ý tên màu chuẩn khi user gõ Tab ở tham số màu
" của :ExcelSetBg/:ExcelSetFg.
" ----------------------------------------------------------------------------
function! ExcelColorComplete(A, L, P) abort
    return filter(copy(s:excel_color_names), 'v:val =~? "^" . a:A')
endfunction
" ============================================================================
" USER-FACING
" ============================================================================
" ----------------------------------------------------------------------------
" Các lệnh thủ công (command) để người dùng gọi trực tiếp nếu cần:
"   :ExcelSave                  -> lưu buffer hiện tại ra file .xlsx
"   :ExcelSheets                 -> liệt kê tên các sheet
"   :ExcelSheetRename            -> đổi tên 1 sheet (qua menu chọn)
"   :ExcelSheetOpen <name>       -> mở/chuyển sang sheet <name> (hỗ trợ Tab complete)
"   :ExcelSheetAdd <name>        -> tạo sheet mới tên <name>
"   :ExcelSheetDelete <name>     -> xoá sheet <name> (hỗ trợ Tab complete)
"   :ExcelSetBg <color>          -> đổi màu nền cell tại con trỏ
"   :'<,'>ExcelSetBg <color>     -> đổi màu nền toàn bộ cell trong vùng chọn (Visual mode)
"   :ExcelSetFg <color>          -> đổi màu chữ cell tại con trỏ (hỗ trợ Visual mode như trên)
"   :ExcelBold                   -> đảo in đậm cell tại con trỏ (hỗ trợ Visual mode)
"   :ExcelItalic                 -> đảo in nghiêng cell tại con trỏ (hỗ trợ Visual mode)
"   :'<,'>ExcelMerge             -> GỘP các ô trong vùng Visual selection thành 1 ô
"                                  (yêu cầu Visual mode, ít nhất 2 ô; giá trị chỉ giữ
"                                  ở ô top-left, các ô khác bị xoá value — giống Excel)
"   :ExcelUnmerge                -> BỎ GỘP vùng chứa ô tại con trỏ
"   :'<,'>ExcelUnmerge           -> bỏ gộp mọi vùng giao với Visual selection
"   :ExcelGoto <ref>             -> nhảy con trỏ đến ô (vd :ExcelGoto B3); Tab để gợi ý
"
"   STATUSLINE: tự động hiển thị "<sheet> / <cell_ref>" ở thanh status —
"   thông tin này KHÔNG nằm trong buffer text nên không bị yank/copy theo,
"   thay thế tốt cho việc hiển thị tiêu đề cột A/B/C/... trong text.
"
"   Màu hợp lệ: red, green, blue, yellow, orange, purple, gray, white, black,
"   "#RRGGBB", hoặc "none" để xoá màu.
"   LƯU Ý: để chọn 1 vùng ô hình chữ nhật giống Excel (ví dụ B2:D5), nên
"   dùng Visual BLOCK mode (Ctrl-V rồi di chuyển) thay vì Visual thường (v)
"   — vì bảng hiển thị là text ASCII, Visual ký tự thường (v) trên dòng đầu
"   và dòng cuối sẽ chọn theo cột thay vì theo ô, có thể cuốn thêm ô không
"   mong muốn ở 2 đầu dòng. Visual Block (Ctrl-V) hoặc Visual dòng (V) đều
"   cho kết quả đúng như chọn vùng trong Excel.
" ----------------------------------------------------------------------------
command! ExcelBuild call ExcelBuild()
command! ExcelSave call ExcelSave()
command! ExcelSheets call ExcelSheets()
command! ExcelSheetRename call ExcelSheetRename()
command! -nargs=1 -complete=customlist,ExcelSheetComplete ExcelSheetOpen call ExcelSheetOpen(<q-args>)
command! -nargs=1 ExcelSheetAdd call ExcelSheetAdd(<q-args>)
command! -nargs=1 -complete=customlist,ExcelSheetComplete ExcelSheetDelete call ExcelSheetDelete(<q-args>)
command! -range -nargs=1 -complete=customlist,ExcelColorComplete ExcelSetBg call ExcelSetBg(<q-args>, <line1>, <line2>)
command! -range -nargs=1 -complete=customlist,ExcelColorComplete ExcelSetFg call ExcelSetFg(<q-args>, <line1>, <line2>)
command! -range ExcelBold call ExcelBold(<line1>, <line2>)
command! -range ExcelItalic call ExcelItalic(<line1>, <line2>)
command! -range ExcelMerge call ExcelMerge(<line1>, <line2>)
command! -range ExcelUnmerge call ExcelUnmerge(<line1>, <line2>)
command! ExcelShowFormula call ExcelShowFormula()
command! -range ExcelApplyFormula call ExcelApplyFormula(<line1>, <line2>)
command! -nargs=1 -complete=customlist,ExcelGotoComplete ExcelGoto call ExcelGoto(<q-args>)
