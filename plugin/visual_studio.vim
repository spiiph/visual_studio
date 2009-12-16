" visual_studio.vim - Visual Studio integration with Vim 
" Author:               Henrik Öhman <speeph@gmail.com>
" URL:                  http://github.com/spiiph/visual_studio/tree/master
" Version:              2.0
" LastChanged:          $LastChangeDate$
" Revision:             $Revision$
" Original Author:      Michael Graz <mgraz.vim@plan10.com>
" Original Copyright:   Copyright (c) 2003-2007 Michael Graz

" Pre-fork history {{{1
" Visual Studio .NET integration with Vim 
"
" Copyright (c) 2003-2007 Michael Graz
" mgraz.vim@plan10.com
"
" Version 1.2 Sep-07
" Support for multiple instances and startup projects
" Thanks for the work of Henrik Öhman <spiiph@hotmail.com>
"
" Version 1.1 May-04
" Support for compiling & building
" Thanks to Leif Wickland for contributing to this feature
"
" Version 1.0 Dec-03
" Base functionality

"----------------------------------------------------------------------
" TODO list {{{1
" - Consider removing support for python exe. Practically all Vim executables
"   for Windows are built with +python anyway
" * Some :echo messages are not displayed during function execution;
"   investigate, and fix if possible.
" - Rename s:visual_studio_xyz to s:xyz - no need for a prefix here.
" * Break out commands, mappings and menu to its own file(s)
" * Change to an autoload-structure for functions

" Load guards {{{1
if exists('loaded_visual_studio')
    finish
endif
"let loaded_visual_studio = 1

" Only run on win32 and win64, not cygwin
if !has("win32") && !has("win64")
    finish
endif

if version < 700
    echomsg "visual_studio.vim plugin requires Vim 7 or above"
    finish
endif

if !has("python")
    echomsg "visual_studio.vim plugin requires a Vim compiled with Python support, " . 
        \ "and that the correct version of Python is installed"
    finish
endif

scriptencoding utf-8

"----------------------------------------------------------------------
" Variables {{{1

" If setting special versions of the following vs_ files,
" make sure to escape backslashes.

"----------------------------------------------------------------------
" InitVariable() function {{{2
" This function is used to initialise a given variable to a given value. The
" variable is only initialised if it does not exist prior
" (Shamelessly stolen from NERD_tree.vim)
function! s:InitVariable(var, value)
    if !exists(a:var)
        if type(a:value) == type([]) || type(a:value) == type({})
            exec 'let ' . a:var . ' = ' . string(a:value)
        else
            exec 'let ' . a:var . ' = ' . "'" . a:value . "'"
        endif

        return 1
    endif
    return 0
endfunction

"----------------------------------------------------------------------
" Global variables {{{2
call s:InitVariable("g:visual_studio_use_location_list", 0)
call s:InitVariable("g:visual_studio_quickfix_height", 20)
call s:InitVariable("g:visual_studio_errorformat", {})
" NOTE: we could include linker errors if we want, but it's fairly useless
"       \'%*\\d>c1xx\ :\ fatal\ error\ %t%n:\ %m:\ ''%f''%.%#' " c1xx errors
call s:InitVariable("g:visual_studio_errorformat['cpp']",
    \ '%*\\d>%f(%l)\ :\ %m,' .
    \ '%f(%l)\ :\ %m,' .
    \ '\ %#%f(%l)\ :\ %m')
call s:InitVariable("g:visual_studio_errorformat['csharp']",
    \ '\ %#%f(%l\\\,%c):\ %m')
call s:InitVariable("g:visual_studio_errorformat['find_results']",
    \ "\ %#%f(%l):%m")
call s:InitVariable("g:visual_studio_errorformat_task_list",
    \ "%f(%l)\ %#:\ %#%m")
call s:InitVariable("g:visual_studio_write_before_build", 1)
call s:InitVariable("g:visual_studio_ignore_file_types",  "obj,lib,res")
call s:InitVariable("g:visual_studio_menu", 1)
call s:InitVariable("g:visual_studio_project_submenus", 1)
call s:InitVariable("g:visual_studio_commands", 1)
call s:InitVariable("g:visual_studio_mappings", 1)

"----------------------------------------------------------------------
" Local variables {{{2
call s:InitVariable("s:module", expand("<sfile>:t:r"))
call s:InitVariable("s:python_init", 0)
call s:InitVariable("s:location", expand("<sfile>:h"))
call s:InitVariable("s:solutions", [])
call s:InitVariable("s:projects", [])
call s:InitVariable("s:solution_index", -1)
call s:InitVariable("s:project_index", -1)
call s:InitVariable("s:output", $TEMP . '\vs_output.txt')

"----------------------------------------------------------------------
" Initialization {{{1

"----------------------------------------------------------------------
" Initialization function {{{2
" Import visual_studio.py
function! s:PythonInit()
    if s:python_init
        return
    endif

    python import sys
    exe 'python sys.path.append(r"' . s:location . '")'
    exe 'python import ' . s:module

    let s:python_init = 1
endfunction

"----------------------------------------------------------------------
" Plugin initialization {{{2
call s:PythonInit()


"----------------------------------------------------------------------
" Python module access functions {{{1

"----------------------------------------------------------------------
" Execute a python function {{{2
" Execute a function in the visual_studio.py module with the supplied
" arguments. If Vim was compiled with +python, use the :python command.
" Otherwise, use the system() function to run the python executable, and
" run :execute on the output.
function! s:DTEExec(py_func, ...)
    " All functions except update_solution_list and set_current_dte require a
    " solution to be selected. If no solution is selected, select the default
    " one
    if index(["update_solution_list", "set_current_dte"], a:py_func) == -1
        if !s:SolutionIsSelected()
            return
        endif
    endif

    "Create the argument string
    let arglist = ['"' . a:py_func . '"']
    for arg in a:000
        call add(arglist, '"' . arg . '"')
    endfor
    let pyargs = join(arglist, ",")

    " Call the python function
    exe printf("python %s.dte_execute(%s)", s:module, pyargs)
endfunction


"----------------------------------------------------------------------
" Debug functions {{{1

"----------------------------------------------------------------------
" Reload python module {{{2
" Force a reload of visual_studio.py. Useful when changing visual_studio.py.
function! DTEReload()
    exe "python import " . s:module
    exe "python reload(" . s:module . ")"
    call s:DTEExec("set_current_dte", s:GetSolutionPID())
    echo s:module . ".py is reloaded."
endfunction

"----------------------------------------------------------------------
" Single file operations {{{1

"----------------------------------------------------------------------
" Get current file {{{2
" Get the current file from Visual Studio, and open it in Vim.
function! DTEGetFile()
    if &modified && !&hidden && !&autowriteall
        call s:DTEExec("get_file", "split")
    else
        call s:DTEExec("get_file", "edit")
    endif
endfunction

"----------------------------------------------------------------------
" Put current file {{{2
" Update the current file and send it to Visual Studio.
function! DTEPutFile()
    update
    let filename = escape(expand("%:p"), '\')
    if filename == ""
        echomsg "No file to send to Visual Studio."
        return 0
    endif
    call s:DTEExec("put_file", filename, line("."), col("."))
    return 1
endfunction

"----------------------------------------------------------------------
" Multiple file operations {{{1

"----------------------------------------------------------------------
" List files in project {{{2
" List all files in a Visual Studio project. If the project name is
" unspecified, the Startup Project is used.
function! DTEListFiles(...)
    " Optional args passed in are
    "  a:1 -- project_name - name of the project to fetch files for
    if a:0 >= 1
        call s:DTEGetProjectFiles(a:1)
    else
        call s:DTEGetProjectFiles()
    end

    for f in s:project_files
        echo f
    endfor
    call input("Press <Enter> to continue ...")
endfunction

"----------------------------------------------------------------------
" Get files in project {{{2
" Add all files in a Visual Studio project to the argument list. If the
" project name is unspecified, the Startup Project is used.
function! DTEGetFiles(...)
    " Optional args passed in are
    "  a:1 -- project_name - name of the project to fetch files for
    echo "Retrieving Project Files ..."
    if a:0 >= 1
        call s:DTEGetProjectFiles(a:1)
    else
        call s:DTEGetProjectFiles()
    end

    if len(s:project_files) == 0
        echomsg "No files found in projects"
    else
        echo "Found " . len(s:project_files) . " file(s)"
        exe "silent args " . join(s:project_files, ' ')
    end
endfunction

"----------------------------------------------------------------------
" Project files helper function {{{2
" Helper function to DTEGetFiles and DTEListFiles.
function! s:DTEGetProjectFiles(...)
    " Optional args passed in are
    "  a:1 -- project_name - name of the project to fetch files for
    if a:0 >= 1 | let project_name = a:1 | endif

    " The following call will assign values to
    " s:project_files
    let s:project_files = []
    if exists("project_name")
        call s:DTEExec("update_project_files_list", project_name)
    else
        call s:DTEExec("update_project_files_list")
    endif

    " Filter files with extensions that should be ignored
    let extensions = split(g:visual_studio_ignore_file_types, ",")
    let s:project_files = s:FilterExtensions(
        \ s:project_files,  
        \ extensions)

    let s:project_files = 
        \ map(s:project_files, 'fnamemodify(v:val, ":.")')
endfunction

function! s:FilterExtensions(files, extensions)
    return filter(a:files, 
        \ 'index(a:extensions, matchstr(v:val, "\\.\\zs[^.]\\+$")) == -1')
endfunction

"----------------------------------------------------------------------
" Quickfix functions {{{1
" Functions for handling output from Visual Studio using quickfix or location
" list commands.

"----------------------------------------------------------------------
" Task list {{{2
" Get the task list from Visual Studio 
function! DTETaskList()
    call s:DTEExec("get_task_list", escape(s:output, '\'))
    call s:DTELoadErrorFile("Task List")
    call s:DTEQuickfixOpen()
endfunction

"----------------------------------------------------------------------
" Output {{{2
" Get the output from a build or compilation from Visual Studio
function! DTEOutput()
    call s:DTEExec("get_output", escape(s:output, '\'), "Output")
    call s:DTELoadErrorFile("Output")
    call s:DTEQuickfixOpen()
endfunction

"----------------------------------------------------------------------
" Find results {{{2
" Get find results from Visual Studio
function! DTEFindResults(which)
    if a:which == 1
        call s:DTEExec("get_output", escape(s:output, '\'),
            \ "Find Results 1")
    else
        call s:DTEExec("get_output", escape(s:output, '\'),
            \ "Find Results 2")
    endif
    s:DTELoadErrorFile("Find Results")
    call s:DTEQuickfixOpen()
endfunction

"----------------------------------------------------------------------
" Load error file {{{2
" Load output, task list or find results from Visual Studio into the quickfix
" list or a location list. 
function! s:DTELoadErrorFile(type)
    " save errorformat
    let saveefm = &errorformat

    " set errorformat
    if a:type == "Task List"
        exe "set errorformat=".
            \ g:visual_studio_errorformat["task_list"]
    elseif a:type == "Find Results"
        exe "set errorformat+=".
            \ g:visual_studio_errorformat["find_results"]
    else
        exe "set errorformat =". 
            \ g:visual_studio_errorformat["cpp"]
        exe "set errorformat+=". 
            \ g:visual_studio_errorformat["csharp"]
    endif

    if g:visual_studio_use_location_list
        exe "lgetfile " . s:output
    else
        exe "cgetfile " . s:output
    endif

    " restore errorformat
    let &errorformat = saveefm
endfunction
        
"----------------------------------------------------------------------
" Open error window {{{2
" Open the quickfix or a location list buffer
function! s:DTEQuickfixOpen()
    if g:visual_studio_quickfix_height > 0
        if g:visual_studio_use_location_list
            exe "lopen " . g:visual_studio_quickfix_height
        else
            exe "copen " . g:visual_studio_quickfix_height
        endif
    endif
endfunction

"----------------------------------------------------------------------
" Build functions {{{1

"----------------------------------------------------------------------
" Compile file {{{2
" Compile the current file
function! DTECompileFile()
    if !DTEPutFile()
        return
    endif
    call s:DTEExec("compile_file", escape(s:output, '\'))
    call s:DTELoadErrorFile("Output")
    call s:DTEQuickfixOpen()
    echo "Done compiling."
endfunction

"----------------------------------------------------------------------
" Build project {{{2
" Build a project. If no argument is supplied, build the Startup Project.
function! DTEBuildProject(...)
    if g:visual_studio_write_before_build
        wall
    endif

    if a:0 >= 1
        call s:DTEExec("build_project", escape(s:output, '\'), a:1)
    else
        call s:DTEExec("build_project", escape(s:output, '\'))
    endif
    call s:DTELoadErrorFile("Output")
    call s:DTEQuickfixOpen()
    echo "Done building project."
endfunction

"----------------------------------------------------------------------
" Build solution {{{2
" Build the current solution.
function! DTEBuildSolution()
    if g:visual_studio_write_before_build
        wall
    endif

    call s:DTEExec("build_solution", escape(s:output, '\'))
    call s:DTELoadErrorFile("Output")
    call s:DTEQuickfixOpen()
    echo "Done building solution."
endfunction

"----------------------------------------------------------------------
" Solution functions {{{1

"----------------------------------------------------------------------
" List solutions {{{2
" List all Visual Studio solutions
function! DTEListSolutions()
    " Populate s:solutions
    call s:DTEGetInstances()
    if len(s:solutions) == 0
        echomsg "No Visual Studio instances found"
    else
        for i in range(len(s:solutions))
            let selected = (s:solution_index == i)
            echo s:CreateMenuEntry(selected, " ", i + 1, s:GetSolutionName(i))
        endfor
    endif
endfunc

"----------------------------------------------------------------------
" Select solution {{{2
" Select a solution by name, or list all available solutions and select by
" solution number.
function! DTESelectSolution(...)
    " Populate s:solutions
    call s:DTEGetInstances()
    if len(s:solutions) == 0
        echomsg "No Visual Studio instances found"
        return
    endif

    " Optional args passed in are
    "  a:1 -- solution name or DEFAULT
    if a:0 > 0
        let index = index(map(copy(s:solutions), "v:val[1]"),
            \ a:1)
        if index == -1
            echomsg "Invalid solution name: " . a:1
        endif
    else
        let menu = ["Select solution"]
        for i in range(len(s:solutions))
            let selected = (s:solution_index == i)
            let entry = s:CreateMenuEntry(selected, " ", i + 1,
                \ s:GetSolutionName(i))
            call add(menu, entry)
        endfor
        let index = inputlist(menu) - 1
        redraw
    endif

    if index >= 0
        if !s:SelectSolutionByIndex(index)
            echomsg "Invalid selection: " . index
        else
            echomsg "Connected: " . s:GetSolutionName()
        endif
    endif
endfunc

"----------------------------------------------------------------------
" Select default solution {{{2
" Select the default solution (index 0) if none is selected.
function! s:SolutionIsSelected()

    if s:solution_index != -1
        return 1
    endif

    call s:DTEGetInstances()
    if len(s:solutions) == 0
        echo "No Visual Studio instances found"
        return 0
    endif

    if !s:SelectSolutionByIndex(0)
        echo "Failed to select the default solution."
        return 0
    else
        return 1
    endif
endfunc

"----------------------------------------------------------------------
" Refresh the solution menu {{{2
" Refresh the VisualStudio.Solutions menu, and display the popup menu.
function! s:MenuRefreshSolutions()
    " Populate s:solutions
    call s:DTEGetInstances()
    if len(s:solutions) == 0
        echo "No Visual Studio instances found"
    else
        echo "Found " . len(s:solutions) . " solutions"
        call s:UpdateSolutionMenu()
        popup! VisualStudio.Solutions
    endif
endfunc

"----------------------------------------------------------------------
" Select solution {{{2
" Select a solution by name, or list all available solutions and select by
" solution number.
function! s:MenuSelectSolution(index)
    if !s:SelectSolutionByIndex(a:index)
        echo "Invalid selection: " . a:index
    else
        echo "Connected: " . s:GetSolutionName()
    endif
endfunc

"----------------------------------------------------------------------
" Get Visual Studio instances {{{2
" Find all Visual Studio instances (i.e. all solutions) and add them to the
" solution list.
function! s:DTEGetInstances()
    let s:solutions = []
    " The following call will populate s:solutions
    call s:DTEExec("update_solution_list")
    call s:UpdateSolutionMenu()
endfunction

"----------------------------------------------------------------------
" Select solution by index {{{2
" Select a solution by supplying its index. Update the solution and project
" menus.
function! s:SelectSolutionByIndex(index)
    if a:index < 0 || a:index >= len(s:solutions)
        return 0
    endif
    "if a:index != s:solution_index
        let s:solution_index = a:index
        call s:DTEExec("set_current_dte", s:GetSolutionPID())
        call s:UpdateSolutionMenu()
    "endif
    call s:DTEGetProjects()
    return 1
endfunction

"----------------------------------------------------------------------
" Update solution menu {{{2
" Update the solution menu after having selected a different solution or
" after having updated the solution list.
function! s:UpdateSolutionMenu()
    if !g:visual_studio_menu
        return
    endif
    try
        aunmenu VisualStudio.Solutions
    catch
    endtry
    for i in range(len(s:solutions))
        let selected = (s:solution_index == i)
        let item = s:CreateMenuEntry(selected, " &", i + 1,
            \ s:GetSolutionName(i))
        exe "amenu <silent> .810 &VisualStudio.&Solutions." . 
            \ escape(item, '\ .') .
            \ " :call <SID>MenuSelectSolution(" . i . ")<CR>"
    endfor
    if len(s:solutions) > 0
        amenu <silent> &VisualStudio.&Solutions.-separator- :
    endif
    amenu <silent> &VisualStudio.&Solutions.&Refresh
        \ :call <SID>MenuRefreshSolutions()<CR>
endfunction

"----------------------------------------------------------------------
" Get solution PID {{{2
" Get the PID of the solution with the supplied index. If index is
" unspecified, get the PID of the current solution.
function! s:GetSolutionPID(...)
    if a:0 > 0 
        let index = a:1
    else
        let index = s:solution_index
    endif
    let item = get(s:solutions, index, [0, ""])
    return item[0]
endfunction

"----------------------------------------------------------------------
" Get solution name {{{2
" Get the name of the solution with the supplied index. If index is
" unspecified, get the name of the current solution.
function! s:GetSolutionName(...)
    if a:0 > 0 
        let index = a:1
    else
        let index = s:solution_index
    endif
    let item = get(s:solutions, index, [0, ""])
    return item[1]
endfunction

"----------------------------------------------------------------------
" Solution completion {{{2
" Command line completion on solution name; return a list of solution names.
function! s:CompleteSolution(ArgLead, CmdLine, CursorPos)
    let result = map(copy(s:solutions), "v:val[1]")
    return result
endfunction

"----------------------------------------------------------------------
" Project functions {{{1

"----------------------------------------------------------------------
" List projects {{{2
" List all projects in the current solution
function! DTEListProjects()
    " Populate s:projects
    if len(s:projects) == 0
        echo "No projects found in solution" 
    else
        for i in range(len(s:projects))
            let selected = (s:project_index == i)
            echo s:CreateMenuEntry(selected, " ", i + 1,
                \ s:GetProjectName(i))
        endfor
    endif
endfunc

"----------------------------------------------------------------------
" Select project {{{2
" Select a project by name, or list all projects in the current solution and
" select by project number.
function! DTESelectProject(...)
    if len(s:projects) == 0
        echo "No projects found in solution" 
        return
    endif

    " Optional args passed in are
    "  a:1 -- project name
    if a:0 > 0 
        let index = index(map(copy(s:projects), "v:val[0]"),
            \ a:1)
        if index == -1
            echo "Invalid project name: " . a:1
        endif
    else
        let menu = ["Select project"]
        for i in range(len(s:projects))
            let selected = (s:project_index == i)
            let entry = s:CreateMenuEntry(selected, " ", i + 1,
                \ s:GetProjectName(i))
            call add(menu, entry)
        endfor
        let index = inputlist(menu) - 1
    endif

    if index >= 0 
        if !s:SelectProjectByIndex(index)
            echo "Invalid selection: " . index
        else
            echo "Selected project: " . s:GetProjectName()
        endif
    endif
endfunc

"----------------------------------------------------------------------
" Refresh the project menu {{{2
" Refresh the VisualStudio.Projects menu for the current solution,
" and display the popup menu.
function! s:MenuRefreshProjects()
    if len(s:projects) == 0
        echo "No projects found in solution"
    else
        echo "Found " . len(s:projects) . " projects"
        call s:UpdateProjectMenu()
        popup! VisualStudio.Projects
    endif
endfunc

"----------------------------------------------------------------------
" Select project {{{2
" Select a solution by name, or list all available solutions and select by
" solution number.
function! s:MenuSelectProject(index)
    if !s:SelectProjectByIndex(a:index)
        echo "Invalid selection: " . a:index
    else
        echo "Set startup project: " . s:GetProjectName()
    endif
endfunc

"----------------------------------------------------------------------
" Get Visual Studio projects {{{2
" Get the projects in the current solution and add them to the project list.
function! s:DTEGetProjects()
    let s:projects = []
    " The following call will populate s:projects
    call s:DTEExec("update_project_list")
    call s:UpdateProjectMenu()
endfunction

"----------------------------------------------------------------------
" Select Visual Studio startup project {{{2
" Select a Visual Studio startup project by supplying its index. Update the
" project menu.
function! s:SelectProjectByIndex(index)
    if a:index < 0 || a:index >= len(s:projects)
        return 0
    endif
    if a:index != s:project_index
        call s:DTEExec("set_startup_project", s:GetProjectName(a:index))
        let s:project_index = a:index
        call s:UpdateProjectMenu()
    endif
    return 1
endfunction

"----------------------------------------------------------------------
" Update project menu {{{2
" Update the project menu after having selected a different project or
" after having updated the project list.
function! s:UpdateProjectMenu()
    if !g:visual_studio_menu
        return
    endif

    try
        aunmenu VisualStudio.Projects
    catch
    endtry

    for i in range(len(s:projects))
        let selected = (s:project_index == i)
        let item = "&VisualStudio.Pro&jects." .
            \ escape(s:CreateMenuEntry(selected, " &", i + 1,
            \ s:GetProjectName(i)), " .")

        exe "amenu <silent> .810 " . item . ".&Build\\ Project" .
            \ " :call DTEBuildProject(<SID>GetProjectName(" . i . "))<CR>"
        exe "amenu <silent> .810 " . item . ".Set\\ Start&up\\ Project" .
            \ " :call <SID>MenuSelectProject(" . i . ")<CR>"
        exe "amenu <silent> .810 " . item . ".&List\\ Project\\ Files " .
            \ ":call DTEListFiles(<SID>GetProjectName(" . i . "))<CR>"
        exe "amenu <silent> .810 " . item . ".&Get\\ Project\\ Files " .
            \ ":call DTEGetFiles(<SID>GetProjectName(" . i . "))<CR>"

        if g:visual_studio_project_submenus
            exe "amenu <silent> .810 " . item . ".-separator- :"

            if type(s:GetProjectChildren(i)) == type([])
                for child in s:GetProjectChildren(i)
                    call s:UpdateProjectSubMenu(item, child[0], child[1])
                endfor
            endif
        endif
    endfor

    if len(s:projects) > 0
        amenu <silent> .810 &VisualStudio.Pro&jects.-separator- :
    endif
    amenu <silent> .810 &VisualStudio.Pro&jects.&Refresh
        \ :call <SID>MenuRefreshProjects()<CR>
endfunction

" Create the project sub menus
function! s:UpdateProjectSubMenu(parent, name, value)
    let item = a:parent . "." . escape(a:name, " .")
    if type(a:value) == type([])
        for child in a:value
            call s:UpdateProjectSubMenu(item, child[0], child[1])
        endfor
    else
        exe "amenu <silent> " . item . 
            \ " :call <SID>OpenProjectSubMenu('" . a:value . "')<CR>"
        return
    endif
endfunction

"----------------------------------------------------------------------
" Get project name {{{2
" Get the name of the project with the supplied index. If index is
" unspecified, get the name of the current project.
function! s:GetProjectName(...)
    if a:0 > 0 
        let index = a:1
    else
        let index = s:project_index
    endif
    let item = get(s:projects, index, ["", []])
    return item[0]
endfunction

"----------------------------------------------------------------------
" Project completion {{{2
" Command line completion on project name; return a list of project names.
function! s:CompleteProject(ArgLead, CmdLine, CursorPos)
    let result = map(copy(s:projects), "v:val[0]")
    return result
endfunction

"----------------------------------------------------------------------
" Get project children {{{2
" Get the children of the project with the supplied index. If index is
" unspecified, get the children of the current project.
function! s:GetProjectChildren(...)
    if a:0 > 0 
        let index = a:1
    else
        let index = s:project_index
    endif
    let item = get(s:projects, index, ["", []])
    return item[1]
endfunction

"----------------------------------------------------------------------
" Open project sub-menu {{{2
" Open a file a the project sub-menu
function! s:OpenProjectSubMenu(filename)
    if &modified && !&hidden && !&autowriteall
        exe "split " . a:filename
    else
        exe "edit " . a:filename
    endif
endfunction

"----------------------------------------------------------------------
" Helper functions {{{1

"----------------------------------------------------------------------
" Create menu entry {{{2
" Helper functino to create a menu item based on selection, a prefix an item
" number and an item string. For both solutions and projects.
function! s:CreateMenuEntry(selected, prefix, number, item)
    if a:selected
        return printf("*%s%d %s", a:prefix, a:number, a:item)
    else
        return printf(" %s%d %s", a:prefix, a:number, a:item)
    endif
endfunction

"----------------------------------------------------------------------
" Help and about {{{1
function! DTEOnline()
    call system("cmd /c start http://www.plan10.com/vim/visual-studio/doc/1.2")
endfunction

function! DTEAbout()
    echo "visual_studio.vim version 1.2+"
    echo "Customizations by Henrik Öhman <speeph@gmail.com>"
    echo "git clone git://github.com/spiiph/visual_studio/tree/master"
    call input("Press <Enter> to continue ...")
endfunction

"----------------------------------------------------------------------
" Mappings, commands and menus {{{1

"----------------------------------------------------------------------
" Menu setup {{{2

if has("gui") && g:visual_studio_menu
    amenu <silent> &VisualStudio.&Get\ File :call DTEGetFile()<CR>
    amenu <silent> &VisualStudio.&Put\ File :call DTEPutFile()<CR>
    amenu <silent> &VisualStudio.&List\ Startup\ Project\ Files
        \ :call DTEListFiles()<CR>
    amenu <silent> &VisualStudio.Get\ Start&up\ Project\ Files
        \ :call DTEGetFiles()<CR>
    amenu <silent> &VisualStudio.-separator1- :<CR>
    amenu <silent> &VisualStudio.&Task\ List :call DTETaskList()<CR>
    amenu <silent> &VisualStudio.&Output :call DTEOutput()<CR>
    amenu <silent> &VisualStudio.&Find\ Results\ 1 :call DTEFindResults(1)<CR>
    amenu <silent> &VisualStudio.Find\ Results\ &2 :call DTEFindResults(2)<CR>
    amenu <silent> &VisualStudio.-separator2- :<CR>
    amenu <silent> &VisualStudio.&Build\ Solution :call DTEBuildSolution()<CR>
    amenu <silent> &VisualStudio.Build\ Start&up\ Project
        \ :call DTEBuildProject()<CR>
    amenu <silent> &VisualStudio.&Compile\ File :call DTECompileFile()<CR>
    amenu <silent> &VisualStudio.-separator3- :<CR>
    call s:UpdateSolutionMenu()
    call s:UpdateProjectMenu()
    amenu <silent> .900 &VisualStudio.-separator4- :<CR>
    amenu <silent> .910 &VisualStudio.&Help.&Online :call DTEOnline()<CR>
    amenu <silent> .910 &VisualStudio.&Help.&About :call DTEAbout()<CR>
endif

"----------------------------------------------------------------------
" Plugin mappings {{{2
nnoremap <silent> <Plug>VSGetFile :call DTEGetFile()<CR>
nnoremap <silent> <Plug>VSPutFile :call DTEPutFile()<CR>
nnoremap <silent> <Plug>VSTaskList :call DTETaskList()<CR>
nnoremap <silent> <Plug>VSOutput :call DTEOutput()<CR>
nnoremap <silent> <Plug>VSFindResults1 :call DTEFindResults(1)<CR>
nnoremap <silent> <Plug>VSFindResults2 :call DTEFindResults(2)<CR>
nnoremap <silent> <Plug>VSBuildSolution :call DTEBuildSolution()<CR>
nnoremap <silent> <Plug>VSBuildProject :call DTEBuildProject()<CR>
nnoremap <silent> <Plug>VSCompileFile :call DTECompileFile()<CR>
nnoremap <silent> <Plug>VSSelectSolution :call DTESelectSolution()<CR>
nnoremap <silent> <Plug>VSSelectProject :call DTESelectProject()<CR>
nnoremap <silent> <Plug>VSListFiles :call DTEListFiles()<CR>
nnoremap <silent> <Plug>VSGetFiles :call DTEGetFiles()<CR>
nnoremap <silent> <Plug>VSAbout :call DTEAbout()<CR>
nnoremap <silent> <Plug>VSOnline :call DTEOnline()<CR>

"----------------------------------------------------------------------
" Default mappings {{{2
if g:visual_studio_mappings
    nmap <silent> <Leader>vg <Plug>VSGetFile
    nmap <silent> <Leader>vp <Plug>VSPutFile
    nmap <silent> <Leader>vt <Plug>VSTaskList
    nmap <silent> <Leader>vo <Plug>VSOutput
    nmap <silent> <Leader>vf <Plug>VSFindResults1
    nmap <silent> <Leader>v2 <Plug>VSFindResults2
    nmap <silent> <Leader>vb <Plug>VSBuildSolution
    nmap <silent> <Leader>vu <Plug>VSBuildProject
    nmap <silent> <Leader>vc <Plug>VSCompileFile
    nmap <silent> <Leader>vs <Plug>VSSelectSolution
    nmap <silent> <Leader>vj <Plug>VSSelectProject
    nmap <silent> <Leader>vl <Plug>VSListFiles
    nmap <silent> <Leader>ve <Plug>VSGetFiles
    nmap <silent> <Leader>va <Plug>VSAbout
    nmap <silent> <Leader>vh <Plug>VSOnline
endif

"----------------------------------------------------------------------
" Command setup {{{2
if g:visual_studio_commands
    com! DTEGetFile call DTEGetFile()
    com! DTEPutFile call DTEPutFile()
    com! DTEOutput call DTEOutput()
    com! DTEFindResults1 call DTEFindResults(1)
    com! DTEFindResults2 call DTEFindResults(2)
    com! DTEBuildSolution call DTEBuildSolution()
    com! -nargs=* -complete=customlist,s:CompleteProject
        \ DTEBuildProject call DTEBuildProject(<f-args>)
    com! -nargs=* -complete=customlist,s:CompleteProject
        \ DTEListFiles call DTEListFiles(<f-args>)
    com! -nargs=* -complete=customlist,s:CompleteProject
        \ DTEGetFiles call DTEGetFiles(<f-args>)
    com! DTECompileFile call DTECompileFile()
    com! -nargs=* -complete=customlist,s:CompleteSolution
        \ DTESelectSolution call DTESelectSolution(<f-args>)
    com! DTEListSolutions call DTEListSolutions()
    com! -nargs=* -complete=customlist,s:CompleteProject
        \ DTESelectProject call DTESelectProject(<f-args>)
    com! DTEListProjects call DTEListProjects()
    com! DTEAbout call DTEAbout()
    com! DTEOnline call DTEOnline()
    com! DTEReload call DTEReload()
endif

" vim: set sts=4 sw=4 fdm=marker:
