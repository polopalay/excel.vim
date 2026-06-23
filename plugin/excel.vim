" ============================================================================
" Plugin: ExcelPlugin
" Mục đích: Cho phép mở/sửa/lưu file .xlsx trực tiếp trong Vim dưới dạng
"           văn bản dạng bảng (text-table), thông qua script Python ở src/excel.py
" ============================================================================

" Ngăn plugin load lại nhiều lần (guard chuẩn của mọi vim plugin)
if exists('g:loaded_excelPlugin')
    finish
endif
let g:loaded_excelPlugin = 1

" Xác định đường dẫn gốc của plugin (lùi lên 1 cấp từ thư mục chứa file này)
" Ví dụ: nếu file này nằm ở plugin/excelPlugin.vim thì root_dir là thư mục cha
let s:root_dir = fnamemodify(expand('<sfile>:p:h'), ':h')
let s:excel_py = s:root_dir . '/src/excel.py'

" Kiểm tra python3 hoặc python có sẵn trong PATH không, ưu tiên python3
if executable('python3')
    let s:python = 'python3'
elseif executable('python')
    let s:python = 'python'
else
    echoerr 'Python not found'
    finish
endif

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
	highlight ExcelUpper  gui=bold cterm=bold
	highlight ExcelItalic gui=italic cterm=italic

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

	" --- Chữ HOA liên tiếp (ví dụ tiêu đề cột, header) -> in đậm ---
	" Hỗ trợ cả ký tự có dấu tiếng Việt (À-Ỹ, Đ)
	call matchadd(
       		\ 'ExcelUpper',
        	\ '\v<[A-ZÀ-ỸĐ]{2,}>',
        	\ 25
        	\ )

	" --- Chữ trong dấu ngoặc đơn (ví dụ chú thích, ghi chú) -> in nghiêng ---
	call matchadd(
            \ 'ExcelItalic',
            \ '\v\([^)]*\)',
            \ 30
            \ )

endfunction

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
"
" Cách làm: kiểm tra xem augroup 'zip' có đang autocmd riêng cho *.xlsx không,
" nếu có thì xoá CHỈ đúng phần *.xlsx đó đi (không đụng đến *.zip, *.jar, ...
" hay autocmd của plugin/augroup khác)
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

function! s:ExcelCmd(mode, ...) abort
    let l:cmd =
                \ s:python . ' '
                \ . shellescape(s:excel_py)
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
" ExcelOpen(): đọc file .xlsx qua script Python, hiển thị nội dung dạng bảng
" text trong buffer hiện tại
" ----------------------------------------------------------------------------
function! ExcelOpen()

    " Tránh gọi đệ quy (vì hàm này có thể tự gọi lại buffer khi reload)
    if exists('b:xlsx_buffer')
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

    " Kiểm tra script Python xử lý Excel có tồn tại không
    if !filereadable(s:excel_py)
        echoerr 'excel.py not found: ' . s:excel_py
        unlet! b:xlsx_buffer
        return
    endif

    " Gọi script Python với lệnh 'open' để chuyển .xlsx -> text dạng bảng
	let l:output = s:ExcelCmd(
            \ 'open',
            \ b:xlsx_sheet
            \ )

    " Nếu script Python lỗi (exit code != 0) thì báo lỗi và dừng
    if v:shell_error
        echoerr join(l:output, "\n")
        unlet! b:xlsx_buffer
        return
    endif

    " Cho phép sửa buffer tạm thời để nạp nội dung mới vào
    setlocal modifiable
    silent %delete _

    " Đưa nội dung bảng (output từ Python) vào buffer, bắt đầu từ dòng 1
    call setline(1, l:output)

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

    " Áp dụng tô màu cú pháp cho nội dung bảng
    call ExcelSetupHighlight()

    " Đăng ký autocmd lưu file: khi user gõ :w, gọi ExcelSave() thay vì
    " để Vim ghi buffer ra file theo cách thông thường
    augroup ExcelBuffer
        autocmd! * <buffer>
        autocmd BufWriteCmd <buffer> call ExcelSave()
    augroup END

    " Đánh dấu buffer là "chưa chỉnh sửa" (vừa mở xong, chưa có thay đổi gì)
    set nomodified

endfunction

" ----------------------------------------------------------------------------
" ExcelSave(): lưu nội dung buffer (dạng bảng text) trở lại thành file .xlsx
" thông qua script Python, sau đó load lại để hiển thị bản đã format chuẩn
" ----------------------------------------------------------------------------
function! ExcelSave()

    " Tạo file tạm để chứa nội dung hiện tại của buffer
    let l:tmp = tempname()

    call writefile(
                \ getline(1, '$'),
                \ l:tmp
                \ )

    " Gọi script Python với lệnh 'save': đọc file tạm (dạng bảng text)
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
	" với nội dung đã được Python format/chuẩn hoá lại (căn cột, v.v.)
	let l:output = s:ExcelCmd(
            \ 'open',
            \ b:xlsx_sheet
            \ )

    if v:shell_error
        echoerr join(l:output, "\n")
        return
    endif

    " Nạp lại nội dung mới (đã format) vào buffer
    setlocal modifiable

    silent %delete _

    call setline(1, l:output)
	call ExcelSetupHighlight()

    " Đánh dấu lại buffer là "đã lưu, không còn thay đổi"
    set nomodified

    echo 'Excel saved & reformatted'

endfunction

function! ExcelSheets()
    echo join(
                \ s:ExcelCmd('sheets'),
                \ "\n"
                \ )
endfunction

function! ExcelOpenSheet(sheet)

    if &modified
        set nomodified
    endif

    let b:xlsx_sheet = a:sheet

    let l:output = s:ExcelCmd(
                \ 'open',
                \ a:sheet
                \ )

    if v:shell_error
        echoerr join(l:output, "\n")
        return
    endif

    setlocal modifiable

    silent %delete _

    call setline(1, l:output)

    call ExcelSetupHighlight()

    set nomodified

endfunction

" ----------------------------------------------------------------------------
" Tự động gán filetype=excel cho mọi file .xlsx, dù mở mới hay đọc lại
" (bổ sung phòng trường hợp filetype chưa được set qua ExcelOpen)
" ----------------------------------------------------------------------------
augroup ExcelFiletype
    autocmd!
    autocmd BufRead,BufNewFile *.xlsx setfiletype excel
augroup END

" ----------------------------------------------------------------------------
" Các lệnh thủ công (command) để người dùng gọi trực tiếp nếu cần:
" :ExcelSave    -> lưu buffer hiện tại ra file .xlsx
" ----------------------------------------------------------------------------
command! ExcelSave call ExcelSave()
command! ExcelSheets call ExcelSheets()
command! -nargs=1 ExcelOpenSheet call ExcelOpenSheet(<q-args>)
