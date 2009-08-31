" Visual Studio .NET integration with Vim
"
" Copyright (c) 2003-2007 Michael Graz
" mgraz.vim@plan10.com
"
" Version 1.2 Sep-07
" Support for multiple instances and startup projects
" Thanks for the work of Henrik �hman <spiiph@hotmail.com>
"
" Version 1.1 May-04
" Support for compiling & building
" Thanks to Leif Wickland for contributing to this feature
"
" Version 1.0 Dec-03
" Base functionality

"if exists('loaded_plugin_visual_studio')
"    finish
"endif
"let loaded_plugin_visual_studio = 1

if version < 700
    echo "visual_studio.vim plugin requires Vim 7 or above"
    finish
endif

"----------------------------------------------------------------------
" Global variables

" If setting special versions of the following vs_ files,
" make sure to escape backslashes.

if ! exists ('g:visual_studio_task_list')
    let g:visual_studio_task_list = $TEMP.'\\vs_task_list.txt'
endif
if ! exists ('g:visual_studio_output')
    let g:visual_studio_output = $TEMP.'\\vs_output.txt'
endif
if ! exists ('g:visual_studio_find_results_1')
    let g:visual_studio_find_results_1 = $TEMP.'\\vs_find_results_1.txt'
endif
if ! exists ('g:visual_studio_find_results_2')
    let g:visual_studio_find_results_2 = $TEMP.'\\vs_find_results_2.txt'
endif
if ! exists ('g:visual_studio_quickfix_height')
    let g:visual_studio_quickfix_height = 20
endif
if ! exists ('g:visual_studio_quickfix_errorformat_cpp')
    let g:visual_studio_quickfix_errorformat_cpp  = 
        \ '%*\\d>%f(%l)\ :\ %m,' .
        \ '%f(%l)\ :\ %m,' .
        \ '\ %#%f(%l)\ :\ %m'
    " we could include linker errors if we want, but it's fairly useless
    " \'%*\\d>c1xx\ :\ fatal\ error\ %t%n:\ %m:\ ''%f''%.%#' " c1xx errors
endif
if ! exists ('g:visual_studio_quickfix_errorformat_csharp')
    let g:visual_studio_quickfix_errorformat_csharp = '\ %#%f(%l\\\,%c):\ %m'
endif
if ! exists ('g:visual_studio_quickfix_errorformat_find_results')
    let g:visual_studio_quickfix_errorformat_find_results = '\ %#%f(%l):%m'
endif
" efm should be modifiable on the fly, and not only when 
" the script is loaded
"if ! exists ('g:visual_studio_quickfix_errorformat')
"    let s:qf1 = g:visual_studio_quickfix_errorformat_cpp
"    let s:qf2 = g:visual_studio_quickfix_errorformat_csharp
"    let s:qf3 = g:visual_studio_quickfix_errorformat_find_results
"    let g:visual_studio_quickfix_errorformat = s:qf1.',\\'.s:qf2.',\\'.s:qf3
"endif
if ! exists ('g:visual_studio_quickfix_errorformat_task_list')
    let g:visual_studio_quickfix_errorformat_task_list = '%f(%l)\ %#:\ %#%m'
endif
if ! exists ('g:visual_studio_quickfix_path')
    let g:visual_studio_quickfix_path = ''
endif
if ! exists('g:visual_studio_quickfix_spec')
    let g:visual_studio_quickfix_spec = 'copen'
endif
if ! exists ('g:visual_studio_has_python')
    let g:visual_studio_has_python = has('python')
endif
if ! exists ('g:visual_studio_python_exe')
    let g:visual_studio_python_exe = 'python.exe'
endif
if ! exists ('g:visual_studio_write_before_build')
    let g:visual_studio_write_before_build = 1
endif
if ! exists('g:visual_studio_ignore_file_types')
    let g:visual_studio_ignore_file_types = "obj,lib"
endif

"----------------------------------------------------------------------
" Local variables

let s:visual_studio_module = 'visual_studio'
let s:visual_studio_python_init = 0
let s:visual_studio_location = expand("<sfile>:h")
let s:visual_studio_pid = 0
let s:visual_studio_lst_dte = []
let s:visual_studio_lst_project = []
let s:visual_studio_startup_project_index = -1

"----------------------------------------------------------------------

function! <Sid>PythonInit()
    if s:visual_studio_python_init
        return 1
    endif
    if g:visual_studio_has_python
        python import sys
        exe 'python sys.path.append(r"'.s:visual_studio_location.'")'
        exe 'python import '.s:visual_studio_module
    else
        call <Sid>PythonDllCheck()
        if ! <Sid>PythonExeCheck()
            return 0
        endif
        let s:visual_studio_module = '"'.s:visual_studio_location.'\'.s:visual_studio_module.'.py"'
    endif
    let s:visual_studio_python_init = 1
    return 1
endfunction

"----------------------------------------------------------------------

function! <Sid>PythonDllCheck()
    redir => output
    silent version
    redir end
    let srch = 'DYNAMIC_PYTHON_DLL='
    let idx = stridx(output, srch)
    if idx <= 1
        echo "\rVim does not appear to be built with a python DLL. Using a version of vim built with embedded python will result in better performance of the visual_studio.vim plugin."
        call input ('Type return to continue...')
        return
    endif
    let dll_name = output[idx+len(srch):]
    let dll_name = substitute (dll_name, ' .*$', '', '')
    let dll_name = substitute (dll_name, '["\\]', '', 'g')
    echo "\rVim has been been built with python DLL '".dll_name."' but it could not be loaded.  If this version of python is available to vim the result will be better performance of the visual_studio.vim plugin."
    call input ('Type return to continue...')
endfunction

"----------------------------------------------------------------------

function! <Sid>PythonExeCheck()
    let output = system (g:visual_studio_python_exe.' -c "print 123"')
    if output !~? '^123'
        echo "\rERROR cannot run: ".g:visual_studio_python_exe
        echo "\rUpdate the system PATH or else set g:visual_studio_python_exe to a valid python.exe"
        return 0
    else
        return 1
    endif
endfunction

"----------------------------------------------------------------------

function! <Sid>GetPid()
    try
        let pid = libcallnr ('msvcrt.dll', '_getpid', '')
        return pid
    catch
    endtry
    try
        let pid = libcallnr ('crtdll.dll', '_getpid', '')
        return pid
    catch
    endtry
    return 0
endfunction

"----------------------------------------------------------------------

function! <Sid>DTEExec(fcn_py, ...)
    if ! <Sid>PythonInit()
        return
    endif

    " NOTE: rewrote argument creation to use join; looks nicer
    let arglist = [s:visual_studio_pid]
    for pyarg in a:000
        let arglist += ["\"" . pyarg . "\""]
    endfor

    if g:visual_studio_has_python
        let pyargs = join(arglist, ',')
        exe 'python ' . s:visual_studio_module . '.' . a:fcn_py . '(' . pyargs . ')'
    else
        let pyargs = join(arglist, ' ')
        let vim_pid = <Sid>GetPid()
        if vim_pid > 0
            let pyargs = pyargs . ' vim_pid=' . vim_pid
        endif
        " Here the output of the python script is executed as a vim command
        let cmd = g:visual_studio_python_exe.' '.s:visual_studio_module.' '.a:fcn_py.' '.pyargs
        let result = system(cmd)
        exe result
    endif
endfunction

"----------------------------------------------------------------------

" NOTE: new function; useful for debugging visual_studio.py
function! DTEReload()
    if ! <Sid>PythonInit()
        return
    endif
    if g:visual_studio_has_python
        exe 'python reload ('.s:visual_studio_module.')'
        exe 'python import '.s:visual_studio_module
        echo s:visual_studio_module . ".py is reloaded."
    endif
endfunction

"----------------------------------------------------------------------

function! DTEGetFile()
    call <Sid>DTEExec ('dte_get_file', &modified)
endfunction

"----------------------------------------------------------------------

function! DTEPutFile()
    let filename = escape(expand('%:p'), '\')
    if filename == ''
        echo 'No vim file!'
        return 0
    endif
    call <Sid>DTEExec ('dte_put_file', filename, &modified, line('.'), col('.'))
    return 1
endfunction

"----------------------------------------------------------------------

function! DTEListFiles(...)
    " Optional args passed in are
    "  a:1 -- project_name - name of the project to fetch files for
    if a:0 >= 1
        call <Sid>DTEGetProjectFiles(a:1)
    else
        call <Sid>DTEGetProjectFiles()
    end

    for f in s:visual_studio_lst_project_files
        echomsg f
    endfor
endfunction

function! DTEGetFiles(...)
    " Optional args passed in are
    "  a:1 -- project_name - name of the project to fetch files for
    echo 'Retrieving Project Files ...'
    if a:0 >= 1
        call <Sid>DTEGetProjectFiles(a:1)
    else
        call <Sid>DTEGetProjectFiles()
    end

    if len(s:visual_studio_lst_project_files) == 0
        echo 'No projects found'
    else
        echomsg 'Found ' . len(s:visual_studio_lst_project_files) . ' file(s)'
        exe "silent args " . join(s:visual_studio_lst_project_files, ' ')
    end
endfunction

function! <Sid>DTEGetProjectFiles(...)
    " Optional args passed in are
    "  a:1 -- project_name - name of the project to fetch files for
    if a:0 >= 1 | let project_name = a:1 | endif

    " The following call will assign values to
    " s:visual_studio_lst_project_files
    let s:visual_studio_lst_project_files = []
    if exists("project_name")
        call <Sid>DTEExec ('dte_get_files', project_name)
    else
        call <Sid>DTEExec ('dte_get_files')
    endif

    " Filter files with extensions that should be ignored
    let extensions = split(g:visual_studio_ignore_file_types, ",")
    let s:visual_studio_lst_project_files = <Sid>FilterExtensions(
        \ s:visual_studio_lst_project_files,  
        \ extensions)

    let s:visual_studio_lst_project_files = 
        \ map(s:visual_studio_lst_project_files, 'fnamemodify(v:val, ":.")')
endfunction

function! <Sid>FilterExtensions(files, extensions)
    return filter(a:files, 
        \ 'index(a:extensions, matchstr(v:val, "\\.\\zs[^.]\\+$")) == -1')
endfunction

"----------------------------------------------------------------------

function! DTETaskList()
    let &errorfile = g:visual_studio_task_list
    call <Sid>DTEExec ('dte_task_list', &errorfile)
endfunction

"----------------------------------------------------------------------

function! DTEOutput()
    let &errorfile = g:visual_studio_output
    call <Sid>DTEExec ('dte_output', &errorfile, 'Output')
endfunction

"----------------------------------------------------------------------

function! DTEFindResults(which)
    if a:which == 1
        let &errorfile = g:visual_studio_find_results_1
        let window_caption = 'Find Results 1'
    else
        let &errorfile = g:visual_studio_find_results_2
        let window_caption = 'Find Results 2'
    endif
    call <Sid>DTEExec ('dte_output', &errorfile, window_caption)
endfunction

"----------------------------------------------------------------------

function! <Sid>DTEQuickfixOpen(which)
    " save errorformat
    let saveefm = &errorformat

    " set errorformat and path
    if a:which == 'Task List'
        exe 'set errorformat='.
            \ g:visual_studio_quickfix_errorformat_task_list
    else
        exe 'set errorformat ='. 
            \ g:visual_studio_quickfix_errorformat_cpp
        exe 'set errorformat+='. 
            \ g:visual_studio_quickfix_errorformat_csharp
        exe 'set errorformat+='.
            \ g:visual_studio_quickfix_errorformat_find_results
    endif

    " open error file
    if g:visual_studio_quickfix_path != ''
        " This is a workaround for VS8 which uses relative files in its error
        " messages. If someone comes up with a better way of doing this, none
        " would be happier than me. (Henrik)
        exe 'silent! cd ' .  g:visual_studio_quickfix_path 
        cfile
        cd -
    else
        cfile
    end

    " open error window
    if g:visual_studio_quickfix_height > 0
        exe g:visual_studio_quickfix_spec . ' ' . g:visual_studio_quickfix_height
    endif

    "restore errorformat
    let &errorformat = saveefm
endfunction

"----------------------------------------------------------------------

function! DTECompileFile()
    if ! DTEPutFile()
        return
    endif
    " NOTE: there used to be a cd (change directory) here but not anymore
    let &errorfile = g:visual_studio_output
    call <Sid>DTEExec ('dte_compile_file', &errorfile)
endfunction

"----------------------------------------------------------------------

function! DTEBuildSolution()
    let &errorfile = g:visual_studio_output
    call <Sid>DTEExec ('dte_build_solution', &errorfile,
        \ g:visual_studio_write_before_build)
endfunction

"----------------------------------------------------------------------

function! DTEBuildProject(...)
    " Optional args passed in are
    "  a:1 -- name of the project to build
    if a:0 >= 1 | let project_name = a:1 | endif

    let &errorfile = g:visual_studio_output
    if exists("project_name")
        call <Sid>DTEExec ('dte_build_project', &errorfile,
            \ g:visual_studio_write_before_build, project_name)
    else
        call <Sid>DTEExec ('dte_build_project', &errorfile,
            \ g:visual_studio_write_before_build)
    endif
endfunction

"----------------------------------------------------------------------

function! DTEGetSolutions(...)
    " Optional args passed in are
    "  a:1 -- if true then gui menu is displayed
    let gui_menu = 0
    if a:0 >= 1 | let gui_menu = a:1 | endif
    let s:visual_studio_lst_dte = []
    " The following call will assign values to s:visual_studio_lst_dte
    echo 'Searching for Visual Studio instances ...'
    call <Sid>DTEExec ('dte_list_instances')
    call <Sid>DTESolutionGuiMenuCreate()
    if len(s:visual_studio_lst_dte) == 0
        echo 'No VisualStudio instances found'
        return
    endif
    if gui_menu
        echo 'Found '.len(s:visual_studio_lst_dte).' VisualStudio instance(s)'
        popup! VisualStudio.Solutions
    else
        call <Sid>DTESolutionTextMenu()
    endif
endfunction

function! <Sid>DTESolutionGuiMenuCreate()
    try
        aunmenu VisualStudio.Solutions
    catch
    endtry
    let menu_num = 0
    for dte in s:visual_studio_lst_dte
        let menu_num = menu_num + 1
        let dte_pid = dte[0]
        let dte_sln = escape (dte[1], '\ .')
        if dte_pid == s:visual_studio_pid
            let leader = '*\ &'.menu_num
        else
            let leader = '\ \ \ &'.menu_num
        endif
        exe 'amenu <silent> .810 &VisualStudio.&Solutions.'.leader.'\ '.dte_sln.' :call <Sid>DTESolutionMenuChoice('.menu_num.')<cr>'
    endfor
    if len(s:visual_studio_lst_dte) > 0
        amenu <silent> &VisualStudio.&Solutions.-separator- :
    endif
    amenu <silent> &VisualStudio.&Solutions.&Refresh :call DTEGetSolutions(1)<cr>
endfunction

function! <Sid>DTESolutionTextMenu()
    if len(s:visual_studio_lst_dte) == 0
        return
    endif
    let menu_num = 0
    let lst_menu = ['Select Solution']
    for dte in s:visual_studio_lst_dte
        let menu_num = menu_num + 1
        let dte_pid = dte[0]
        let dte_sln = dte[1]
        if dte_pid == s:visual_studio_pid
            let leader = '* '.menu_num
        else
            let leader = '  '.menu_num
        endif
        let lst_menu += [leader.' '.dte_sln]
    endfor
    let choice = inputlist(lst_menu)
    if choice > 0
        if ! <Sid>DTESolutionMenuChoice(choice)
            echo 'Invalid selection: '.choice
        endif
    endif
endfunction

function! <Sid>DTESolutionMenuChoice(which)
    if a:which <= 0 || a:which > len(s:visual_studio_lst_dte)
        return 0
    endif
    let index = a:which - 1
    let s:visual_studio_pid = s:visual_studio_lst_dte[index][0]
    echo 'Connected: '.s:visual_studio_lst_dte[index][1]
    call <Sid>DTESolutionGuiMenuCreate()
    call DTEGetProjects(0)
    return 1
endfunction

"----------------------------------------------------------------------

" Get the projects in the current solution, add them to
" visual_studio_lst_project and show the menu.
function! DTEGetProjects(...)
    " Optional args passed in are
    "  a:1 -- verbose  - if true then give feedback re projects
    "  a:2 -- gui_menu - if true then gui menus displayed
    let verbose = 1
    let gui_menu = 0
    if a:0 >= 1 | let verbose = a:1 | endif
    if a:0 >= 2 | let gui_menu = a:2 | endif
    let s:visual_studio_lst_project = []

    " The following call will assign values to s:visual_studio_lst_project
    if verbose
        echo 'Retrieving Projects ...'
    endif
    call <Sid>DTEExec ('dte_list_projects')
    call <Sid>DTEProjectGuiMenuCreate()
    if len(s:visual_studio_lst_project) == 0
        if verbose
            echo 'No projects found'
        endif
        return
    endif
    if ! verbose
        return
    endif
    if gui_menu
        echo 'Found '.len(s:visual_studio_lst_project).' project(s)'
        popup! VisualStudio.Projects
    else
        call <Sid>DTEProjectTextMenu()
    endif
endfunction

" Set the startup project by name and update
" visual_studio_startup_project_index accordingly
function! DTESetStartupProject(name)
    if <Sid>DTEExec ('dte_set_startup_project', a:name)
        " set the current startup project index variable
        s:visual_studio_startup_project_index = 
            \ index(s:visual_studio_lst_project, a:name)
        echo "Startup project set to " . a:name
        " update the project menu
        call <Sid>DTEProjectGuiMenuCreate()
    endif
endfunction

" Create the Projects menu
function! <Sid>DTEProjectGuiMenuCreate()
    try
        aunmenu VisualStudio.Projects
    catch
    endtry
    let menu_num = 0
    for project in s:visual_studio_lst_project
        let project_name = escape(project[0], '\. ')
        let project_children = project[1]
        let menu_num = menu_num + 1
        if menu_num == s:visual_studio_startup_project_index + 1
            let leader = '*\ &'.menu_num
        else
            let leader = '\ \ \ &'.menu_num
        endif
        let project_name_menu = '&VisualStudio.Pro&jects.' . leader . 
            \ '\ ' . project_name

        " set up the project menus
        exe 'amenu <silent> .820 ' . project_name_menu . '.&Build\ Project ' .
            \ ' :call DTEBuildProject("' . project_name . '")<cr>'
        exe 'amenu <silent> .820 ' . project_name_menu . '.Set\ Start&up\ Project ' .
            \ ':call DTESetStartupProject("' . project_name . '")<cr>'
        exe 'amenu <silent> .820 ' . project_name_menu . '.&List\ Project\ Files ' .
            \ ':call DTEListProjectFiles("' . project_name . '")<cr>'
        exe 'amenu <silent> .820 ' . project_name_menu . '.&Get\ Project\ Files ' .
            \ ':call DTEGetProjectFiles("' . project_name . '")<cr>'
        exe 'amenu <silent> .820 '.project_name_menu.'.-separator- :<cr>'

        " set up the menus for project children
        if type(project_children) == type([])
            for child in project_children
                call <Sid>DTEProjectGuiSubMenuCreate(project_name_menu, child[0], child[1])
            endfor
        endif
    endfor
    if len(s:visual_studio_lst_project) > 0
        amenu <silent> .820 &VisualStudio.Pro&jects.-separator- :
    endif
    amenu <silent> .820 &VisualStudio.Pro&jects.&Refresh :call DTEGetProjects(1, 1)<cr>
endfunction

" Create the project sub menus
function! <Sid>DTEProjectGuiSubMenuCreate(menu, name, value)
    " menu will be quoted, name and value will be unquoted
    let label = a:menu.'.'.escape(a:name, ' .')
    if type(a:value) == type('')
        exe 'amenu <silent> ' . label . 
            \ ' :call <Sid>DTEProjectGuiSubMenuChoice("' . a:value . '")<cr>'
        return
    endif
    for child in a:value
        call <Sid>DTEProjectGuiSubMenuCreate(label, child[0], child[1])
    endfor
endfunction

" Helper for the project sub menu; open the child of the project as a file
function! <Sid>DTEProjectGuiSubMenuChoice(filename)
    if &modified
        exe 'split '.a:filename
    else
        exe 'edit '.a:filename
    endif
endfunction

" Display the project selection text menu
function! <Sid>DTEProjectTextMenu()
    if len(s:visual_studio_lst_project) == 0
        return
    endif
    let menu_num = 0
    let lst_menu = ['Select Startup Project']
    for project in s:visual_studio_lst_project
        let project_name = project[0]
        let menu_num = menu_num + 1
        if menu_num == s:visual_studio_startup_project_index + 1
            let leader = '* '.menu_num
        else
            let leader = '  '.menu_num
        endif
        let lst_menu += [leader.' '.project_name]
    endfor
    let choice = inputlist(lst_menu)
    if choice > 0
        if ! <Sid>DTEProjectMenuStartupChoice(choice)
            echo 'Invalid selection: '.choice
        endif
    endif
endfunction

function! <Sid>DTEProjectMenuStartupChoice(which)
    if a:which <= 0 || a:which > len(s:visual_studio_lst_project)
        return 0
    endif
    let project_index = a:which - 1
    let project_name = s:visual_studio_lst_project[project_index][0]
    call <Sid>DTEExec ('dte_set_startup_project', project_name, project_index)
    call <Sid>DTEProjectGuiMenuCreate()
    return 1
endfunction

"----------------------------------------------------------------------

function! DTEOnline()
    call system('cmd/c start http://www.plan10.com/vim/visual-studio/doc/1.2')
endfunction

function! DTEAbout()
    echo 'visual_studio.vim version 1.2'
endfunction

"----------------------------------------------------------------------
" Menu setup

if has('gui') && ( ! exists('g:visual_studio_menu') || g:visual_studio_menu != 0 )
    amenu <silent> &VisualStudio.&Get\ File :call DTEGetFile()<cr>
    amenu <silent> &VisualStudio.&Put\ File :call DTEPutFile()<cr>
    amenu <silent> &VisualStudio.&List\ Startup\ Project\ Files :call DTEListFiles()
    amenu <silent> &VisualStudio.Get\ Start&up\ Project\ Files :call DTEGetFiles()
    amenu <silent> &VisualStudio.-separator1- :
    amenu <silent> &VisualStudio.&Task\ List :call DTETaskList()<cr>
    amenu <silent> &VisualStudio.&Output :call DTEOutput()<cr>
    amenu <silent> &VisualStudio.&Find\ Results\ 1 :call DTEFindResults(1)<cr>
    amenu <silent> &VisualStudio.Find\ Results\ &2 :call DTEFindResults(2)<cr>
    amenu <silent> &VisualStudio.-separator2- :
    amenu <silent> &VisualStudio.&Build\ Solution :call DTEBuildSolution()<cr>
    amenu <silent> &VisualStudio.Build\ Start&up\ Project :call DTEBuildProject()<cr>
    amenu <silent> &VisualStudio.&Compile\ File :call DTECompileFile()<cr>
    amenu <silent> &VisualStudio.-separator3- :
    call <Sid>DTESolutionGuiMenuCreate()
    call <Sid>DTEProjectGuiMenuCreate()
    amenu <silent> .900 &VisualStudio.-separator4- :
    amenu <silent> .910 &VisualStudio.&Help.&Online :call DTEOnline()<cr>
    amenu <silent> .910 &VisualStudio.&Help.&About :call DTEAbout()<cr>
endif

"----------------------------------------------------------------------
" Mapping setup

" plugin mappings
nnoremap <silent> <Plug>VSGetFile :call DTEGetFile()<CR>
nnoremap <silent> <Plug>VSPutFile :call DTEPutFile()<CR>
nnoremap <silent> <Plug>VSTaskList :call DTETaskList()<CR>
nnoremap <silent> <Plug>VSOutput :call DTEOutput()<CR>
nnoremap <silent> <Plug>VSFindResults1 :call DTEFindResults(1)<CR>
nnoremap <silent> <Plug>VSFindResults2 :call DTEFindResults(2)<CR>
nnoremap <silent> <Plug>VSBuildSolution :call DTEBuildSolution()<CR>
nnoremap <silent> <Plug>VSBuildProject :call DTEBuildProject()<CR>
nnoremap <silent> <Plug>VSCompileFile :call DTECompileFile()<CR>
nnoremap <silent> <Plug>VSGetSolutions :call DTEGetSolutions()<CR>
nnoremap <silent> <Plug>VSGetProjects :call DTEGetProjects()<CR>
nnoremap <silent> <Plug>VSListFiles :call DTEListFiles()<CR>
nnoremap <silent> <Plug>VSGetFiles :call DTEGetFiles()<CR>
nnoremap <silent> <Plug>VSAbout :call DTEAbout()<CR>
nnoremap <silent> <Plug>VSOnline :call DTEOnline()<CR>

if !exists ('g:visual_studio_mapping') || ! g:visual_studio_mapping
    nmap <silent> <Leader>vg <Plug>VSGetFile
    nmap <silent> <Leader>vp <Plug>VSPutFile
    nmap <silent> <Leader>vt <Plug>VSTaskList
    nmap <silent> <Leader>vo <Plug>VSOutput
    nmap <silent> <Leader>vf <Plug>VSFindResults1
    nmap <silent> <Leader>v2 <Plug>VSFindResults2
    nmap <silent> <Leader>vb <Plug>VSBuildSolution
    nmap <silent> <Leader>vu <Plug>VSBuildProject
    nmap <silent> <Leader>vc <Plug>VSCompileFile
    nmap <silent> <Leader>vs <Plug>VSGetSolutions
    " TODO: vju, vjb options ??
    nmap <silent> <Leader>vj <Plug>VSGetProjects
    nmap <silent> <Leader>vl <Plug>VSListFiles
    nmap <silent> <Leader>ve <Plug>VSGetFiles
    nmap <silent> <Leader>va <Plug>VSAbout
    nmap <silent> <Leader>vh <Plug>VSOnline
endif

"----------------------------------------------------------------------
" Command setup
if ! exists ('g:visual_studio_commands') || g:visual_studio_commands != 0
    com! DTEGetFile call DTEGetFile()
    com! DTEPutFile call DTEPutFile()
    com! DTEOutput call DTEOutput()
    com! DTEFindResults1 call DTEFindResults(1)
    com! DTEFindResults2 call DTEFindResults(2)
    com! DTEBuildSolution call DTEBuildSolution()
    com! -nargs=* DTEBuildProject call DTEBuildProject(<f-args>)
    com! -nargs=* DTEListFiles call DTEListFiles(<f-args>)
    com! -nargs=* DTEGetFiles call DTEGetFiles(<f-args>)
    com! DTECompileFile call DTECompileFile()
    com! DTEGetSolutions call DTEGetSolutions()
    com! DTEGetProjects call DTEGetProjects()
    com! DTEAbout call DTEAbout()
    com! DTEOnline call DTEOnline()
    com! DTEReload call DTEReload()
endif

" vim: set sts=4 sw=4:
