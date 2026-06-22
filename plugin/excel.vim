if exists('g:loaded_excelPlugin')
    finish
endif
let g:loaded_excelPlugin = 1

let s:root_dir = fnamemodify(expand('<sfile>:p:h'), ':h')
let s:excel_py = s:root_dir . '/src/excel.py'

if executable('python3')
    let s:python = 'python3'
elseif executable('python')
    let s:python = 'python'
else
    echoerr 'Python not found'
    finish
endif

" Syntax highlighting for Excel files
function! ExcelSetupHighlight()
    silent! call clearmatches()
    " Define highlight groups
    highlight ExcelBorder guifg=#9ca0b0 ctermfg=245
    highlight ExcelDate   guifg=#40a02b ctermfg=70
    highlight ExcelLink   guifg=#0000ff ctermfg=12 gui=underline cterm=underline
	highlight ExcelUpper  gui=bold cterm=bold
	highlight ExcelItalic gui=italic cterm=italic

    highlight default link ExcelNumber Number

    " Borders
    call matchadd('ExcelBorder', '[+]', 5)
    call matchadd('ExcelBorder', '[|]', 5)
    call matchadd('ExcelBorder', '-\+', 5)

    " URLs
    call matchadd(
            \ 'ExcelLink',
            \ '\v%(https?|ftp)://[[:alnum:]/._~:?#[\]@!$&''()*+,;=%-]+',
            \ 30
            \ )

    " Numbers
    call matchadd(
                \ 'ExcelNumber',
                \ '\v<\d+(\.\d+)?>',
                \ 15
                \ )

    " Dates
    call matchadd(
            \ 'ExcelDate',
            \ '\v<('
            \ . '\d{1,2}/\d{1,2}/\d{2,4}(\s+\d{1,2}:\d{2}(:\d{2})?)?'
            \ . '|'
            \ . '\d{1,2}/\d{4}'
            \ . '|'
            \ . '\d{4}-\d{1,2}-\d{1,2}(\s+\d{1,2}:\d{2}(:\d{2})?)?'
            \ . '|'
            \ . '\d{1,2}-\d{1,2}-\d{4}'
            \ . ')>',
            \ 20
            \ )

	" Bold uppercase words (e.g., headers) -> Bold
	call matchadd(
            \ 'ExcelUpper',
            \ '\v<[A-ZÀ-ỸĐ]{2,}(\s+[A-ZÀ-ỸĐ]{2,})+>',
            \ 25
            \ )

	" Italic words (e.g., comments) -> Italic
	call matchadd(
            \ 'ExcelItalic',
            \ '\v\([^)]*\)',
            \ 30
            \ )

endfunction

" When opening an .xlsx file, call the ExcelOpen function
augroup ExcelViewer
    autocmd!
    autocmd BufReadCmd *.xlsx call ExcelOpen()
augroup END

function! ExcelOpen()

    " tránh recurse
    if exists('b:xlsx_buffer')
        return
    endif

    let b:xlsx_buffer = 1

    let l:file = expand('<amatch>')

    if empty(l:file)
        let l:file = expand('%:p')
    endif

    let b:xlsx_file = fnamemodify(l:file, ':p')

    if !filereadable(s:excel_py)
        echoerr 'excel.py not found: ' . s:excel_py
        unlet! b:xlsx_buffer
        return
    endif

    let l:cmd =
                \ s:python . ' '
                \ . shellescape(s:excel_py)
                \ . ' open '
                \ . shellescape(b:xlsx_file)

    let l:output = systemlist(l:cmd)

    if v:shell_error
        echoerr join(l:output, "\n")
        unlet! b:xlsx_buffer
        return
    endif

    setlocal modifiable
    silent %delete _

    call setline(1, l:output)

    setlocal buftype=acwrite
    setlocal bufhidden=hide
    setlocal noswapfile
    setlocal filetype=excel

    call ExcelSetupHighlight()

    augroup ExcelBuffer
        autocmd! * <buffer>
        autocmd BufWriteCmd <buffer> call ExcelSave()
    augroup END

    set nomodified

endfunction

function! ExcelSave()

    let l:tmp = tempname()

    call writefile(
                \ getline(1, '$'),
                \ l:tmp
                \ )

    let l:cmd =
                \ s:python . ' '
                \ . shellescape(s:excel_py)
                \ . ' save '
                \ . shellescape(b:xlsx_file)
                \ . ' '
                \ . shellescape(l:tmp)

    let l:result = systemlist(l:cmd)
	call delete(l:tmp)

    if v:shell_error
        echoerr join(l:result, "\n")
        return
    endif

	" Reopen the file to update the buffer with the new content

    let l:reload_cmd =
                \ s:python . ' '
                \ . shellescape(s:excel_py)
                \ . ' open '
                \ . shellescape(b:xlsx_file)

    let l:output = systemlist(l:reload_cmd)

    if v:shell_error
        echoerr join(l:output, "\n")
        return
    endif

    setlocal modifiable

    silent %delete _

    call setline(1, l:output)
	call ExcelSetupHighlight()

    set nomodified

    echo 'Excel saved & reformatted'

endfunction

function! ExcelCurrentRow()

    let l:row = line('.')

    " chỉ cho phép đứng trên dòng dữ liệu
    if getline('.') !~ '^|'
        return -1
    endif

    return ((l:row - 2) / 2) + 1

endfunction

function! ExcelInsertRowBelow()

    let l:row = ExcelCurrentRow()

    if l:row < 1
        echo "Select a data row"
        return
    endif

    let l:cmd =
                \ s:python . ' '
                \ . shellescape(s:excel_py)
                \ . ' insert_row '
                \ . shellescape(b:xlsx_file)
                \ . ' '
                \ . (l:row + 1)

    call system(l:cmd)
 	unlet! b:xlsx_buffer
    call ExcelOpen()

endfunction

function! ExcelInsertRowAbove()

    let l:row = ExcelCurrentRow()

    if l:row < 1
        echo "Select a data row"
        return
    endif

    let l:cmd =
                \ s:python . ' '
                \ . shellescape(s:excel_py)
                \ . ' insert_row '
                \ . shellescape(b:xlsx_file)
                \ . ' '
                \ . l:row

    call system(l:cmd)

    unlet! b:xlsx_buffer
    call ExcelOpen()

endfunction

function! ExcelCurrentCol()

    let l:line = getline('.')

    if l:line !~ '^|'
        return -1
    endif

    let l:pos = col('.')

    let l:count = 0

    for l:i in range(0, l:pos - 1)
        if l:line[l:i] == '|'
            let l:count += 1
        endif
    endfor

    return l:count

endfunction

function! ExcelInsertColRight()

    let l:col = ExcelCurrentCol()

    if l:col < 1
        echo "Select a cell"
        return
    endif

    let l:cmd =
                \ s:python . ' '
                \ . shellescape(s:excel_py)
                \ . ' insert_col '
                \ . shellescape(b:xlsx_file)
                \ . ' '
                \ . (l:col + 1)

    call system(l:cmd)
 	unlet! b:xlsx_buffer
    call ExcelOpen()

endfunction

function! ExcelInsertColLeft()

    let l:col = ExcelCurrentCol()

    if l:col < 1
        echo "Select a cell"
        return
    endif

    let l:cmd =
                \ s:python . ' '
                \ . shellescape(s:excel_py)
                \ . ' insert_col '
                \ . shellescape(b:xlsx_file)
                \ . ' '
                \ . l:col

    call system(l:cmd)

    unlet! b:xlsx_buffer
    call ExcelOpen()

endfunction

augroup ExcelFiletype
    autocmd!
    autocmd BufRead,BufNewFile *.xlsx setfiletype excel
augroup END

" Automatic commands for manual reload and save
command! ExcelReload call ExcelOpen()
command! ExcelSave call ExcelSave()
