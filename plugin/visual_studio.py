############################################################ {{{1
# Documentation
'''\
visual_studio.py - Companion file to visual_studio.vim
Version: 2.0-beta
Author: Henrik Ohman <speeph@gmail.com>
URL: http://github.com/spiiph/visual_studio
Original Author: Michael Graz <mgraz.vim@plan10.com>
Original Copyright: Copyright (c) 2003-2007 Michael Graz
'''

############################################################ {{{1
# Original work
# See 'visual_studio.vim' for a description of the original work and the
# rationale behind creating this derivative work.

############################################################ {{{1
# TODO list
#  * Exchange VimExt.command with VimExt.set_var for variables.
#  * Investigate the behaviour of DTEWrapper.get_project, and what effect it
#    has on the user experience.

############################################################ {{{1
# Imports
import os
import re
import sys
import time
# NOTE: 'python import pywintypes' fails with PyWin32 builds > 214.
import pywintypes
import pythoncom
import win32com.client


############################################################ {{{1
# Vim module
# NOTE: Vim must be compiled with +python support.
# NOTE: If used as a standalone module, imported directly in Python, the vim
#       module is not available. The VimExt wrappers then prints vim commands
#       and expressions to stdout rather than evaluating them.
try:
    import vim
except ImportError:
    pass

############################################################ {{{1
# Logging initialization
import logging
import tempfile
logger = logging.getLogger('VS')
logger.addHandler(logging.NullHandler())
logger.setLevel(logging.DEBUG)

############################################################ {{{1
class DTEWrapper:
    '''The DTE class encapsulates DTE objects and functionality.'''

    ############################################################ {{{2
    # Initialization and properties
    def __init__(self):
        # Dict containing {pid: dte} pairs
        self.dtes = {}

        # The pid of the current DTE object
        self.current_dte = 0

        # State variable for the UseFullPaths property
        #self.use_full_paths = None

    def __get_dte(self):
        if self.current_dte == 0:
            self.set_current_dte()
        return self.dtes[self.current_dte]
    dte = property(__get_dte)

    def __get_solution(self):
        if self.dte is not None:
            return self.dte.Solution
        else:
            return None
    solution = property(__get_solution)

    def __get_projects(self):
        if self.solution is not None:
            return self.solution.Projects
        else:
            return None
    projects = property(__get_projects)

    def __get_properties(self):
        if self.solution is not None:
            return self.solution.Properties
        else:
            return None
    properties = property(__get_properties)

    def __get_solution_build(self):
        if self.solution is not None:
            return self.solution.SolutionBuild
        else:
            return None
    solution_build = property(__get_solution_build)

    def __get_active_configuration(self):
        if self.solution_build is not None:
            return self.solution_build.ActiveConfiguration
        else:
            return None
    active_configuration = property(__get_active_configuration)


    ############################################################ {{{2
    # Generic helper functions
    def get_project(self, name = None):
        def compare_name(project):
            return project.Name == name

        if name is None:
            name = self.properties.Item("StartupProject").Value

        logger.debug("%s: project name is %s" %
                (func_name(), name))
        logger.debug("%s: available project names %s" %
                (func_name(), [p.Name for p in self.projects]))
        try:
            project = next(
                    p for p in self.projects if compare_name(p))
            return project
        except StopIteration:
            raise

    def get_tools(self, project):
        log_func()

        if self.dte is None:
            return None
        else:
            try:
                return project.Object.Configurations.Item(
                        self.dte.Solution.SolutionBuild.ActiveConfiguration.Name
                        ).Tools
            except AttributeError:
                return None

    def get_compiler_tool(self, project):
        log_func()

        if self.dte is None:
            return None

        tools = self.get_tools(project)
        if tools is None:
            return None
        else:
            return tools.Item("VCCLCompilerTool")

    # Get a property from a property collection; ignore exceptions and return
    # default if one is encountered
    def get_property(self, collection, name, default = None):
        try:
            return collection.Properties.Item(name).Value
        except Exception, e:
            logger.error("Failed to access property %s in collection %s." %
                    (name, collection.Name))
            return default

    # Set a property in a property collection; ignore exceptions
    def set_property(self, collection, name, value):
        try:
            collection.Properties.Item(name).Value = value
            return True
        except Exception, e:
            return False

    # Check if a project item is a file
    def is_file(self, item):
        return item.Kind == u'{6BB5F8EE-4483-11D3-8BCF-00C04F8EC28C}'

    ############################################################ {{{2
    def set_current_dte(self, pid = 0):
        '''Get the DTE object corresponding to pid. If pid is 0, get the
        current DTE object. If the current DTE object is None, get the first
        DTE object in the self.dtes dict.'''

        log_func()

        pid = int(pid)
        if pid == 0 or pid == self.current_dte:
            if (self.current_dte != 0 and
                    self.dtes[self.current_dte] is not None):
                try:
                    self.dtes[self.current_dte].Solution
                    return self.dtes[self.current_dte]
                except Exception, e:
                    logger.exception(e)
                    del self.dtes[self.current_dte]

        try:
            self.update_dtes()
            if self.dtes.has_key(pid):
                self.current_dte = pid
            else:
                self.current_dte = self.dtes.keys()[0]
            return self.dtes[self.current_dte]
        except pywintypes.com_error, e:
            logger.exception(e)
            VimExt.echowarn("Cannot access DTE object. Is Visual Studio running?")
            return None

    ############################################################ {{{2
    def update_dtes(self):
        '''Populate the self.dtes dict with {pid: dte} elements from the
        Running Object Table.'''

        log_func()

        self.dtes = {}
        rot = pythoncom.GetRunningObjectTable()
        rot_enum = rot.EnumRunning()
        context = pythoncom.CreateBindCtx(0)
        while 1:
            monikers = rot_enum.Next()
            if not monikers:
                break

            display_name = monikers[0].GetDisplayName(context, None)
            if display_name.startswith("!VisualStudio.DTE"):
                logger.debug("%s: found instance %s"
                        % (func_name(), display_name))

                try:
                    pid = int(display_name.rpartition(":")[-1])
                except ValueError, e:
                    logger.exception(e)
                    pid = 0

                dte = win32com.client.Dispatch(
                        rot.GetObject(monikers[0]).QueryInterface(
                            pythoncom.IID_IDispatch))
                self.dtes[pid] = dte

    ############################################################ {{{2
    def wait_for_build(self):
        '''Wait for Visual Studio to complete the build.'''

        log_func()

        # Visual Studio build state flags
        vsBuildStateNotStarted = 1
        vsBuildStateInProgress = 2
        vsBuildStateDone = 3

        build = self.solution_build
        try:
            while build.BuildState == vsBuildStateInProgress:
                time.sleep(0.1)
        except Exception, e:
            logger.exception(e)

    ############################################################ {{{2
    def activate(self):
        '''Activate Visual Studio.'''

        log_func()

        if self.dte is None:
            return

        try:
            self.dte.MainWindow.Activate()
            logger.debug("%s: main window caption is %s" %
                    (func_name(), self.dte.MainWindow.Caption))
        except (TypeError, pywintypes.com_error), e:
            logger.error("Failed to activate Visual Studio main window.")

    ############################################################ {{{2
    def set_autoload(self):
        '''Activate the Autoload option in Visual Studio.'''

        log_func()

        if self.dte is None:
            return

        try:
            properties = self.dte.Properties("Environment", "Documents")
            self.set_property(properties, "DetectFileChangesOutsideIDE", 1)
            self.set_property(properties, "AutoloadExternalChanges", 1)
        except pywintypes.com_error, e:
            logger.exception(e)

    ############################################################ {{{2
    def set_use_full_paths(self, project_name = None):
        '''Set the 'Use full Paths' property in the specified project
        or all projects.'''

        log_func()

        if self.dte is None:
            return

        # If project_name is not given, modify all projects
        if project_name is not None:
            projects = [p for p in self.projects if p.Name == project_name]
        else:
            projects = self.projects

        for p in projects:
            compiler = self.get_compiler_tool(p)

            if compiler is not None:
                compiler.UseFullPaths = True
            else:
                logger.debug("%s: compiler is None for project %s" %
                        (func_name(), p.Name))

    ############################################################ {{{2
    def get_task_list(self, output_file):
        '''Retrieves the task list from Visual Studio.'''

        log_func()

        VimExt.set_var("s:command_status", 0)

        if self.dte is None:
            return

        self.dte.ExecuteCommand("View.TaskList")
        task_list_window = None
        for window in self.dte.Windows:
            if str(window.Caption).startswith("Task List"):
                task_list_window = window
        if task_list_window is None:
            VimExt.echowarn("Task List window not active.")
            return

        task_list = task_list_window.Object.TaskItems
        for item in task_list:
            try:
                filename = item.FileName
            except Exception, e:
                logger.exception(e)
                filename = "<no-filename>"

            try:
                line = item.Line
            except Exception, e:
                logger.exception(e)
                line = "<no-line>"

            try:
                description = item.Description
            except Exception, e:
                logger.exception(e)
                description = "<no-description>"

            f.write("%s(%s) : %s" % (filename, line, description))
        f.close()
        VimExt.set_var("s:command_status", 1)

    ############################################################ {{{2
    def get_output(self, output_file, caption):
        '''Fetch the output from a command and write it to a file.'''

        log_func()

        VimExt.set_var("s:command_status", 0)

        if self.dte is None:
            return

        try:
            window = self.dte.Windows.Item(caption)
        except pywintypes.com_error, e:
            logger.exception(e)
            VimExt.echowarn("Window not active (%s)." % caption)
            return

        if caption == "Output":
            sel = window.Object.OutputWindowPanes.Item("Build").\
                TextDocument.Selection
        else:
            sel = window.Selection
        sel.SelectAll()
        f = file(output_file, "w")
        f.write(sel.Text.replace('\r', ''))
        f.close()
        sel.Collapse()

        VimExt.set_var("s:command_status", 1)

    ############################################################ {{{2
    def compile_file(self, output_file):
        '''Compile the current file.'''

        log_func()

        VimExt.set_var("s:command_status", 0)

        if self.dte is None:
            return

        try:
            self.dte.ExecuteCommand("Build.Compile")
            # Wait for build to complete
            self.wait_for_build()
            self.get_output(output_file, "Output")
            VimExt.set_var("s:command_status", 1)
        except Exception, e:
            logger.error("Failed to compile file: %s" % e[2][2])
            VimExt.echowarn("Failed to compile file.")
        VimExt.activate()

    ############################################################ {{{2
    def build_project(self, output_file, project_name = None):
        '''Build a project by name or build the startup project.'''

        log_func()

        VimExt.set_var("s:command_status", 0)

        if self.dte is None:
            return

        self.set_autoload()
        #self.activate()

        try:
            project = self.get_project(project_name)
            config = self.active_configuration.Name
            self.set_use_full_paths(project.Name)

            logger.info("%s: config = %s, unique name = %s" %
                    (func_name(), config, project.UniqueName))
            self.solution_build.BuildProject(config, project.UniqueName, 1)

            # Wait for build to complete
            self.wait_for_build()
            self.get_output(output_file, "Output")
            VimExt.set_var("s:command_status", 1)
        except Exception, e:
            logger.exception(e)
            VimExt.echowarn("Failed to build project.")
        VimExt.activate()

    ############################################################ {{{2
    def build_solution(self, output_file):
        '''Build the solution in the current DTE.'''

        log_func()

        VimExt.set_var("s:command_status", 0)

        if self.dte is None:
            return

        self.set_autoload()
        #self.activate()

        try:
            self.set_use_full_paths()
            self.solution_build.Build(1)
            # Wait for build to complete
            self.wait_for_build()
            self.get_output(output_file, "Output")
            VimExt.set_var("s:command_status", 1)
        except Exception, e:
            logger.exception(e)
            VimExt.echowarn("Failed to build solution.")
        VimExt.activate()

    ############################################################ {{{2
    def set_startup_project(self, project_name):
        '''Set the startup project in Visual Studio.'''

        log_func()

        if self.dte is None:
            return

        if not self.set_property(
                self.solution,
                "StartupProject",
                project_name):
            VimExt.echowarn("Failed to set startup project.")

    ############################################################ {{{2
    def get_file(self, action):
        '''Get the current file from Visual Studio.'''

        log_func()

        if self.dte is None:
            return

        doc = self.dte.ActiveDocument
        if doc is None:
            VimExt.echowarn("No file active in Visual Studio.")
            return
        point = doc.Selection.ActivePoint
        path = os.path.join(doc.Path, doc.Name)
        VimExt.command("%s +%d %s" %(action, point.Line, path))
        VimExt.command("normal %d|" % point.DisplayColumn)

    ############################################################ {{{2
    def put_file(self, filename, line, col):
        '''Send the current file to Visual Studio.'''

        log_func()

        if self.dte is None:
            return

        logger.debug("%s: absolute path %s" %
                (func_name(), os.path.abspath(filename)))

        self.set_autoload()
        item_op = self.dte.ItemOperations.OpenFile(
                os.path.abspath(filename))
        sel = self.dte.ActiveDocument.Selection
        sel.MoveToLineAndOffset(line, col)
        self.activate()

    ############################################################ {{{2
    def update_solution_list(self):
        '''Update Vim's list of solutions.'''

        log_func()

        self.update_dtes()
        instances = [[key, str(self.dtes[key].Solution.FullName)]
                for key in self.dtes.keys()]
        VimExt.command("let s:solutions = %s" % instances)

    ############################################################ {{{2
    def update_project_list(self):
        '''Update Vim's list of projects.'''

        log_func()

        if self.dte is None:
            return

        startup_project_name = self.get_property(self.solution, "StartupProject")
        startup_project_index = -1
        index = 0
        projects = []
        for project in sorted(self.projects, cmp = lambda x,y: cmp(x.Name, y.Name)):
            # Count projects without a Properties object as special projects
            # that shouldn't be listed.
            if project.Properties is None:
                continue
            if project.Name == startup_project_name:
                startup_project_index = index
            projects.append(str(project.Name))
            index += 1
        VimExt.command("let s:projects = %s" % projects)
        VimExt.command("let s:project_index = %s" %
                startup_project_index)

    ############################################################ {{{2
    def update_project_tree(self, project_name = None):
        '''Update Vim's tree of project files for the named project or the
        startup project.'''

        log_func()

        if self.dte is None:
            return

        project_tree = []
        try:
            project = self.get_project(project_name)
            project_tree = self.get_project_tree(project)
        except Exception, e:
            logger.exception(e)
            VimExt.echowarn("Failed to update project tree.")
        VimExt.command("let s:project_tree = %s" % project_tree)

    ############################################################ {{{3
    def get_project_tree(self, item):
        '''Returns a tree (nested lists) of projects and files in projects.
        The first item is the item or item item name. The second item
        contains a list of children or a filename.'''

        log_func()

        if self.is_file(item):
            path = self.get_property(item, "FullPath")
            if path is not None:
                return [str(item.Name), str(path)]

        # Find children to the item
        children = []

        # Item has a sub project
        #if item.SubProject is not None:
        #    children += self.get_project_tree(item.SubProject)

        # Item has project items
        if item.ProjectItems is not None:
            temp = [self.get_project_tree(x) for x in item.ProjectItems]
            logger.debug("%s: found project item name %s" %
                    (func_name(), temp))
            children += temp

        return [str(item.Name), children]

    ############################################################ {{{2
    def update_project_files_list(self, project_name = None):
        '''Update Vim's list of files for the named project or the startup
        project.'''

        log_func()

        if self.dte is None:
            return

        files = []
        try:
            project = self.get_project(project_name)
            files = self.get_project_items_files(project.ProjectItems)
        except Exception, e:
            logger.exception(e)
            VimExt.echowarn("Failed to update project files.")
        VimExt.command("let s:project_files = %s" % files)

    ############################################################ {{{3
    def get_project_items_files(self, items):
        '''Recursive function that returns a list of files in a ProjectItems
        object.'''

        log_func()

        files = []
        for item in items:
            # Add the item itself
            if self.is_file(item):
                path = self.get_property(item, "FullPath")
                if path is not None:
                    files.append(str(path))

            # Add files from subprojects
            #if item.SubProject is not None:
                #sub_project_items = item.SubProject.ProjectItems
                #files += self.get_project_items_files(sub_project_items)

            # Add files from subitems
            files += self.get_project_items_files(item.ProjectItems)

        return files

############################################################ {{{1
class WScriptShell:
    def __init__(self):
        self.__wsh = win32com.client.Dispatch("WScript.Shell")

    def __call__(self, *args, **kw):
        return self.__wsh

############################################################ {{{1
class VimExt:
    ############################################################ {{{2
    # Initialization
    '''Vim extension class for DTEWrapper.'''

    @classmethod
    ############################################################ {{{2
    def get_pid(cls):
        '''Get PID of Vim.'''
        return os.getpid()

    @classmethod
    # NOTE: Does this even work anymore?
    ############################################################ {{{2
    def activate(cls):
        '''Activate Vim.'''
        wsh().AppActivate(VimExt.get_pid())

    @classmethod
    ############################################################ {{{2
    def command(cls, command):
        '''Send an Ex command to Vim using vim.command(). vim.command() is
        wrapped for standalone usage.'''
        command = command.replace("\\\\", "\\")
        log_func()
        if 'vim' in globals():
            vim.command(command)
        else:
            print "Vim command: %s" % command

    @classmethod
    ############################################################ {{{2
    def eval(cls, expr):
        '''Evaluate an expression in Vim using vim.eval(). vim.eval() is
        wrapped for standalone usage.'''
        log_func()
        return vim.eval(expr)

    @classmethod
    ############################################################ {{{2
    def set_var(cls, var, value):
        '''Set a Vim variable to value.'''
        log_func()
        VimExt.command("let %s = %s" % (var, value))

    @classmethod
    ############################################################ {{{2
    def get_var(cls, var):
        '''Get the value of a Vim variable as a string, list, or dict.'''
        log_func()
        if 'vim' in globals():
            return VimExt.eval(var)
        else:
            print "Vim eval: %s" % var
            return 0

    @classmethod
    ############################################################ {{{2
    def exception(cls, e, trace):
        '''Print the output of an exception in Vim.'''
        if isinstance(e, pywintypes.com_error):
            try:
                msg = e[2][2]
            except Exception, e:
                logger.exception(e)
                msg = None
        else:
            msg = e
        if msg is None:
            msg = "Encountered unknown exception"
        VimExt.echoerr(msg)
        VimExt.echomsg("Check %s for details" % log_file)

    @classmethod
    ############################################################ {{{2
    def escape(cls, s):
        '''Escape a string for :echo, :echomsg and :echoerr.'''
        return str(s).replace("'", "''")

    @classmethod
    ############################################################ {{{2
    def echo(cls, msg):
        ''':echo a message.'''
        VimExt.command("echo '%s'" % VimExt.escape(msg))

    @classmethod
    ############################################################ {{{2
    def echomsg(cls, msg):
        ''':echomsg a message.'''
        VimExt.command("echomsg '%s'" % VimExt.escape(msg))

    @classmethod
    ############################################################ {{{2
    def echowarn(cls, msg):
        ''':echomsg with :echohl a message.'''
        VimExt.command("echohl WarningMsg | echomsg 'Error: %s' | echohl None" %
                VimExt.escape(msg))

    @classmethod
    ############################################################ {{{2
    def echoerr(cls, msg):
        ''':echoerr a message.'''
        VimExt.command("echoerr 'Error: %s'" %
                VimExt.escape(msg))

############################################################ {{{1
# Global objects
wsh = WScriptShell()
dte = DTEWrapper()

############################################################ {{{1
# Entry point function
def dte_execute(name, *args):
    '''Wrapper function for calling functions in the global DTE object.'''

    if not hasattr(dte, name):
        VimExt.echoerr("No such function %s." % name)
    else:
        function = getattr(dte, name)
        function(*args)

def dte_cleanup():
    logger.removeHandler(fh)

############################################################ {{{2
# Global helper functions
import inspect
def func_name():
    return inspect.stack()[1][3]

def log_func():
    def kv_to_str(k, v):
        return str(k) + ", " + str(v)

    # get the stack frame of the previous function
    frame = inspect.stack()[1]

    # function name
    func_name = frame[3]

    # convert argument dict to a comma separated string
    arg_info = inspect.getargvalues(frame[0])

    # remove 'self' argument
    if 'self' in arg_info.args:
        arg_info.args.remove('self')
        del arg_info.locals['self']

    # remove 'cls' argument
    if 'cls' in arg_info.args:
        arg_info.args.remove('cls')
        del arg_info.locals['cls']

    # format the args
    args = inspect.formatargvalues(*arg_info)

    logger.info("%s%s" % (func_name, args))

############################################################ {{{1
# Logging setup
log_file = ""
log_level = int(VimExt.get_var("g:visual_studio_log_level"))

if log_level > 0:
    log_file = os.path.join(
            tempfile.gettempdir(),
            "visual_studio_%d.log" % os.getpid())
    fh = logging.FileHandler(log_file)
    formatter = logging.Formatter(
            "%(asctime)s %(levelname)8s: %(message)s",
            "%Y-%m-%d %H:%M:%S")
    fh.setFormatter(formatter)
    logger.addHandler(fh)

    if log_level == 1:
        fh.setLevel(logging.INFO)
    else:
        fh.setLevel(logging.DEBUG)

    logger.info(("Logging started with log level %s") % (
        logging.getLevelName(logger.getEffectiveLevel())))

VimExt.set_var("s:log_file", "'%s'" % log_file)

# vim: set sts=4 sw=4 fdm=marker:
# vim: fdt=v\:folddashes\ .\ "\ "\ .\ substitute(getline(v\:foldstart+1),\ '^\\s\\+\\|#\\s*\\|\:',\ '',\ 'g'):
