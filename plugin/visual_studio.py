############################################################ {{{1
# Documentation
'''\
visual_studio.py - Companion file to visual_studio.vim
Version: 2.0
Author: Henrik Ohman <speeph@gmail.com>
URL: http://github.com/spiiph/visual_studio/tree/master
Original Author: Michael Graz <mgraz.vim@plan10.com>
Original Copyright: Copyright (c) 2003-2007 Michael Graz
'''

############################################################ {{{1
# Pre-fork history
#'''
#Companion file to visual_studio.vim
#Version 1.2
#
#Copyright (c) 2003-2007 Michael Graz
#mgraz.vim@plan10.com
#'''

############################################################ {{{1
# Imports
import os
import re
import sys
import time
import pywintypes
import pythoncom
import win32com.client

############################################################ {{{1
# Vim module
# If Vim was compiled with +python support, import the vim module here.
# Check for 'vim is not None' where we neet do distinguish between using
# vim.command() and print for interaction with Vim.
try:
    import vim
except:
    pass


############################################################ {{{1
# Exit function
# Disable vim exit processing to avoid an exception at exit
#sys.exitfunc = lambda: None

############################################################ {{{1
# Logging
import logging
logging_enabled = False

if logging_enabled:
    import tempfile
    log_file = os.path.join(os.path.dirname(__file__), "visual_studio.log")
    logging.basicConfig(
        filename = log_file,
        filemode = "a",
        level = logging.DEBUG,
        format = "%(asctime)s %(levelname)-8s %(pathname)s(%(lineno)d)\n" +
            "%(message)s\n"
        )
else:
    logging.basicConfig(level = sys.maxint)

logging.info("Started logging at %s" % time.ctime())

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
        self.use_full_paths = None

    def __get_dte(self):
        if self.current_dte == 0:
            self.set_current_dte()
        return self.dtes[self.current_dte]
    dte = property(__get_dte)

    def __get_solution(self):
        if self.dte is not None:
            return self.dte.Solution
    solution = property(__get_solution)

    def __get_projects(self):
        if self.dte is not None:
            return self.dte.Solution.Projects
    projects = property(__get_projects)

    def __get_properties(self):
        if self.dte is not None:
            return self.dte.Solution.Properties
    properties = property(__get_properties)

    def __get_solution_build(self):
        if self.dte is not None:
            return self.dte.Solution.SolutionBuild
    solution_build = property(__get_solution_build)
        
    def __get_active_configuration(self):
        if self.dte is not None:
            return self.dte.Solution.SolutionBuild.ActiveConfiguration
    active_configuration = property(__get_active_configuration)

    ############################################################ {{{2
    def set_current_dte(self, pid = 0):
        '''Get the DTE object corresponding to pid. If pid is 0, get the
        current DTE object. If the current DTE object is None, get the first
        DTE object in the self.dtes dict.'''
        logging.info("== DTE.set_current_dte pid: %s" % str(pid))

        pid = int(pid)

        if pid == 0 or pid == self.current_dte:
            if self.current_dte != 0 and \
                    self.dtes[self.current_dte] is not None:
                try:
                    self.dtes[self.current_dte].Solution
                    return self.dtes[self.current_dte]
                except Exception, e:
                    logging.exception(e)
                    del self.dtes[self.current_dte]

        try:
            self.update_dtes()
            if self.dtes.has_key(pid):
                self.current_dte = pid
            else:
                self.current_dte = self.dtes.keys()[0]
            return self.dtes[self.current_dte]
        except pywintypes.com_error, e:
            logging.exception(e)
            VimExt.echomsg("Error: Cannot access DTE object. " +
                "Is Visual Studio running?")
            return None

    ############################################################ {{{2
    def update_dtes(self):
        '''Populate the self.dtes dict with {pid: dte} elements from the
        Running Object Table.'''
        logging.info("== DTE.update_dtes")
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
                logging.info("== DTE.update_dtes display_name: %s"
                        % display_name)

                try:
                    pid = int(display_name.rpartition(":")[-1])
                except ValueError, e:
                    logging.exception(e)
                    pid = 0

                dte = win32com.client.Dispatch(
                        rot.GetObject(monikers[0]).QueryInterface(
                            pythoncom.IID_IDispatch))
                self.dtes[pid] = dte

    ############################################################ {{{2
    def wait_for_build(self):
        '''Wait for Visual Studio to complete the build.'''
        # Visual Studio build state flags
        vsBuildStateNotStarted = 1
        vsBuildStateInProgress = 2
        vsBuildStateDone = 3

        try:
            while self.solution_build.BuildState == vsBuildStateInProgress:
                time.sleep(0.1)
        except Exception, e:
            logging.exception(e)

    ############################################################ {{{2
    def activate(self):
        '''Activate Visual Studio.'''
        if self.dte is None:
            return

        try:
            logging.info("== DTE.activate caption: %s" %
                    self.dte.MainWindow.Caption)
            self.dte.MainWindow.Activate()
            #wsh().AppActivate(self.dte.MainWindow.Caption)
        except pywintypes.com_error, e:
            logging.exception(e)
            VimExt.echomsg("Error: Cannot access WScript Shell object.")

    ############################################################ {{{2
    def set_autoload(self):
        '''Activate the Autoload option in Visual Studio.'''
        if self.dte is None:
            return

        enabled = []
        try:
            properties = self.dte.Properties("Environment", "Documents")
            for item in ["DetectFileChangesOutsideIDE",
                    "AutoloadExternalChanges"]:
                if not properties.Item(item).Value:
                    p.Item(item).Value = 1
                    enabled.append(item)
        except pywintypes.com_error, e:
            logging.exception(e)
        if len(enabled) != 0:
            VimExt.echo("Enabled %s in VisualStudio" % " and ".join(enabled))

    ############################################################ {{{2
    def set_use_full_paths(self, project_name = None):
        '''Set the 'Use full Paths' property in the specified project
        or all projects.'''

        if self.use_full_paths is not None:
            return

        if self.dte is None:
            return

        self.use_full_paths = {}
        projects = self.projects
        if project_name is not None:
            projects = [p for p in projects if p.Name == project_name]

        for p in projects:
            tool = p.Object.Configurations.Item(
                    self.active_configuration.Name).Tools.Item(
                            "VCCLCompilerTool")
            if tool is not None:
                self.use_full_paths[p.Name] = tool.UseFullPaths
                tool.UseFullPaths = True
            else:
                logging.debug("== DTE.set_use_full_paths tool is " +
                    "None for project: %s" % p.Name)

    ############################################################ {{{2
    def reset_use_full_paths(self, project_name = None):
        '''Reset the 'Use full Paths' property in the specified project
        or all projects.'''
        logging.debug("== DTE.reset_use_full_paths use_full_paths: %s" %
                self.use_full_paths)

        if self.use_full_paths is None:
            return

        if self.dte is None:
            return

        projects = self.projects
        if project_name is not None:
            projects = [p for p in projects if p.Name == project_name]

        for p in projects:
            tool = p.Object.Configurations.Item(
                    self.active_configuration.Name).Tools.Item(
                            "VCCLCompiler")
            if tool is not None:
                tool.UseFullPaths = self.use_full_paths[p.Name]
            else:
                logging.debug("== DTE.set_use_full_paths tool is " +
                    "None for project: %s" % p.Name)
        self.use_full_paths = None

    ############################################################ {{{2
    def has_csharp_projects(self):
        '''Check if a solution has C# projects.'''
        if self.dte is None:
            return

        try:
            if self.dte.CSharpProjects.Count:
                return 1
        except (pywintypes.com_error, AttributeError), e:
            logging.exception(e)
        return 0

    ############################################################ {{{2
    def get_output(self, output_file, caption):
        '''Fetch the output from a command and write it to a file.'''
        logging.info("== DTE.output output_file, caption: %s, %s" %
                (output_file, caption))

        if self.dte is None:
            return

        try:
            window = self.dte.Windows.Item(caption)
        except pywintypes.com_error, e:
            logging.exception(e)
            VimExt.echomsg("Error: window not active (%s)" % caption)
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

    ############################################################ {{{2
    def compile_file(self, output_file):
        '''Compile the current file.'''
        logging.info("== DTE.compile_file output_file: %s" % output_file)

        if self.dte is None:
            return

        try:
            self.dte.ExecuteCommand("Build.Compile")
            # Wait for build to complete
            self.wait_for_build()
            self.get_output(output_file, "Output")
            VimExt.activate()
        except Exception, e:
            logging.exception(e)
            VimExt.exception(e, sys.exc_traceback)
            VimExt.activate()
            return

    ############################################################ {{{2
    def build_project(self, output_file, project_name = None):
        '''Build a project by name or build the startup project.'''
        logging.info("== DTE.build_project output_file, project_name: %s, %s" %
                (output_file, project_name))

        if self.dte is None:
            return

        if self.has_csharp_projects():
            self.dte.Documents.CloseAll()
        self.set_autoload()
        self.activate()

        try:
            if project_name is None:
                project_name = self.properties.Item("StartupProject").Value
            project = [x for x in self.projects
                    if x.Name == project_name][0]
            config = project.Object.Configurations.Item(
                    self.active_configuration.Name)

            logging.info(("== DTE.build_project config, name: %s, %s") %
                    (config, project.UniqueName))

            self.set_use_full_paths(project.Name)
            config.Build()
            # Wait for build to complete
            self.wait_for_build()
            self.get_output(output_file, "Output")
            VimExt.activate()
        except Exception, e:
            logging.exception(e)
            VimExt.exception(e, sys.exc_traceback)
            VimExt.activate()
            return
        finally:
            self.reset_use_full_paths(project.Name)

    ############################################################ {{{2
    def build_solution(self, output_file):
        '''Build the solution in the current DTE.'''
        logging.info("== DTE.build_solution output_file: %s" % output_file)

        if self.dte is None:
            return

        if has_csharp_projects():
            self.dte.Documents.CloseAll()
        self.set_autoload()
        self.activate()

        try:
            self.set_use_full_paths()
            self.solution_build.Build(1)
            # Wait for build to complete
            self.wait_for_build()
            self.get_output(output_file, "Output")
            VimExt.activate()
        except Exception, e:
            logging.exception(e)
            VimExt.exception(e, sys.exc_traceback)
            VimExt.activate()
            return
        finally:
            self.reset_use_full_paths()

    ############################################################ {{{2
    def set_startup_project(self, project_name):
        '''Set the startup project in Visual Studio.'''
        logging.info("== DTE.set_startup_project project_name: %s" %
                project_name)

        if self.dte is None:
            return

        try:
            self.properties.Item("StartupProject").Value = \
                project_name
        except Exception, e:
            logging.exception(e)
            VimExt.exception(e, sys.exc_traceback)
            return
        #VimExt.status("Startup project set to %s" % project_name)

    ############################################################ {{{2
    def get_file(self, action):
        '''Get the current file from Visual Studio.'''
        logging.info("== DTE.get_file action: %s" % action)

        if self.dte is None:
            return

        doc = self.dte.ActiveDocument
        if doc is None:
            VimExt.echomsg("Error: No file active in Visual Studio!")
            return
        point = doc.Selection.ActivePoint
        path = os.path.join(doc.Path, doc.Name)
        commands = [
            "%s +%d %s" %(action, point.Line, path),
            "normal %d|" % point.DisplayColumn,
        ]
        VimExt.command(commands)

    ############################################################ {{{2
    def put_file(self, filename, line, col):
        '''Send the current file to Visual Studio.'''
        logging.info("== DTE.put_file filename, line, col: %s, %s, %s" %
                (filename, line, col))

        logging.info("== DTE.put_file abspath %s" %
                (os.path.abspath(filename)))

        if self.dte is None:
            return

        self.set_autoload()
        item_op = self.dte.ItemOperations.OpenFile(
                os.path.abspath(filename))
        sel = self.dte.ActiveDocument.Selection
        sel.MoveToLineAndOffset(line, col)
        self.activate()

    ############################################################ {{{2
    def update_solution_list(self):
        '''Update Vim's list of solutions.'''
        logging.info("== DTE.update_solution_list")
        self.update_dtes()
        instances = [[key, str(self.dtes[key].Solution.FullName)]
                for key in self.dtes.keys()]
        VimExt.command("let s:visual_studio_solutions = %s" % instances)

    ############################################################ {{{2
    def update_project_list(self):
        '''Update Vim's list of projects.'''
        logging.info("== DTE.update_project_list")

        if self.dte is None:
            return

        startup_project_name = \
            self.properties.Item("StartupProject").Value
        startup_project_index = -1
        index = 0
        projects = []
        for project in sorted(self.projects,
                cmp = lambda x,y: cmp(x.Name, y.Name)):
            # Count projects without a Properties object as special projects
            # that shouldn't be listed.
            if project.Properties is None:
                continue
            if project.Name == startup_project_name:
                startup_project_index = index
            projects.append(self.get_project_tree(project))
            index += 1
        VimExt.command("let s:visual_studio_projects = %s" % projects)
        VimExt.command("let s:visual_studio_project_index = %s" %
                startup_project_index)

    ############################################################ {{{3
    def get_project_tree(self, item):
        '''Returns a tree (nested lists) of projects and files in projects.
        The first item is the item or item item name. The second item
        contains a list of children or a filename.'''
        logging.debug("== get_project_tree name: %s" % item.Name)

        # NOTE: This assumes that only VCFile item kinds are of interest
        #       as leaves, and that VCFile items are always leaves.
        if item.Properties is not None:
            # NOTE: Solution folders don't have the "Kind" property.
            #       Therefore, just ignore exceptions from these and process
            #       their project items below.
            try:
                if item.Properties.Item("Kind").Value == "VCFile":
                    return [str(item.Name),
                            str(item.Properties.Item("FullPath").Value)]
            except pywintypes.com_error, e:
                pass

        # If we didn't find a leaf (VCFile), recurse deeper
        children = []

        # Item has a sub project
        #if item.SubProject is not None:
        #    children += self.get_project_tree(item.SubProject)

        # Item has project items
        if item.ProjectItems is not None:
            temp = [self.get_project_tree(x) for x in item.ProjectItems]
            logging.debug("== get_project_tree name: %s" % temp)
            children += temp

        return [str(item.Name), children]

    ############################################################ {{{2
    def update_project_files_list(self, project_name = None):
        '''Update Vim's list of files for the named project or the startup
        project.'''
        logging.info("== DTE.get_project_files project_name: %s" %
            project_name)

        if self.dte is None:
            return

        files = []
        try:
            if not project_name:
                project_name = self.properties.Item("StartupProject").Value
            project = [x for x in self.projects
                    if x.Name == project_name][0]
            files = self.get_project_items_files(project.ProjectItems)
        except Exception, e:
            logging.exception(e)
            VimExt.exception(e, sys.exc_traceback)
        VimExt.command("let s:visual_studio_project_files = %s" % files)

    ############################################################ {{{3
    def get_project_items_files(self, items):
        '''Recursive function that returns a list of files in a ProjectItems
        object.'''
        files = []
        for i in items:
            # NOTE: This assumes that only VCFile item kinds are of interest
            #       as leaves.
            if i.Properties.Item("Kind").Value == "VCFile":
                files.append(str(i.Properties.Item("FullPath").Value))
            if i.SubProject is not None:
                files += self.get_project_items_files(i.SubProject.ProjectItems)
            files += self.get_project_items_files(i.ProjectItems)
        return files

    ############################################################ {{{2
    def get_task_list(self, output_file):
        logging.info("== DTE.task_list output_file: %s" % output_file)

        if self.dte is None:
            return

        self.dte.ExecuteCommand("View.TaskList")
        task_list_window = None
        for window in self.dte.Windows:
            if str(window.Caption).startswith("Task List"):
                task_list_window = window
        if task_list_window is None:
            VimExt.echomsg("Error: Task List window not active")
            return
        #task_list = task_list_window.Object

        for item in task_list_window.Object.TaskItems:
            try:
                filename = item.FileName
            except Exception, e:
                logging.exception(e)
                filename = "<no-filename>"

            try:
                line = item.Line
            except Exception, e:
                logging.exception(e)
                line = "<no-line>"

            try:
                description = item.Description
            except Exception, e:
                logging.exception(e)
                description = "<no-description>"

            f.write("%s(%s) : %s" % (filename, line, description))
        f.close()

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

    __pid = None
    __has_python = None

    @classmethod
    ############################################################ {{{2
    def set_pid(cls, pid):
        '''Set the PID of Vim when issued from the command line.'''
        VimExt.__pid = pid

    @classmethod
    ############################################################ {{{2
    def has_python(cls):
        '''Check if Vim was compiled with the +python feature.'''
        if VimExt.__has_python is None:
            if globals().get("vim") is None:
                VimExt.__has_python = 0
            else:
                VimExt.__has_python = int(vim.eval("has('python')"))
        return VimExt.__has_python

    @classmethod
    ############################################################ {{{2
    def activate(cls):
        '''Activate Vim.'''
        if VimExt.__pid is None:
            VimExt.__pid = os.getpid()
        wsh().AppActivate(VimExt.__pid)

    @classmethod
    ############################################################ {{{2
    def command(cls, command_list):
        '''Send a Vim command to Vim, either using vim.command() or print, if
        called from the command line.'''
        logging.info("== VimExt.command command_list %s" % command_list)
        if type(command_list) is not type([]):
            command_list = [command_list]
        for cmd in command_list:
            cmd = cmd.replace("\\\\", "\\")
            if VimExt.has_python():
                vim.command(cmd)
            else:
                print cmd

    @classmethod
    ############################################################ {{{2
    def exception(cls, e, trace):
        '''Print the output of an exception in Vim.'''
        if isinstance(e, pywintypes.com_error):
            try:
                msg = e[2][2]
            except Exception, e:
                logging.exception(e)
                msg = None
        else:
            msg = e
        if msg is None:
            msg = "Encountered unknown exception"
        VimExt.echoerr("Error: %s" % msg)
        while trace:
            VimExt.echoerr("    File '%s', line %d, in %s" %
                    (trace.tb_frame.f_code.co_filename, trace.tb_lineno,
                        trace.tb_frame.f_code.co_name))
            trace = trace.tb_next

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
    def echoerr(cls, msg):
        ''':echoerr a message.'''
        VimExt.command("echoerr '%s'" % VimExt.escape(msg))

############################################################ {{{1
# Global objects
wsh = WScriptShell()
dte = DTEWrapper()

############################################################ {{{1
def dte_execute(name, *args):
    '''Wrapper function for calling functions in the global DTE object.'''
    logging.info("== dte_execute name: %s" % name)
    if not hasattr(dte, name):
        VimExt.echoerr("Error: no such function %s" % name)
    else:
        function = getattr(dte, name)
        function(*args)

############################################################ {{{1
def main():
    '''Function handling the indirection when called from the command
    line.'''
    prog = os.path.basename(sys.argv[0])
    logging.info("== main sys.argv: %s" % sys.argv)
    if len(sys.argv) == 1:
        VimExt.echoerr("Error: Not enough arguments to %s" % prog)

    if sys.argv[-1].startswith("vim_pid="):
        VimExt.set_pid(int(sys.argv[-1].rpartition("=")[-1]))
        del sys.argv[-1]

    try:
        dte_execute(sys.argv[1], *sys.argv[2:])
    except TypeError, e:
        logging.exception(e)
        VimExt.exception(e, sys.exc_traceback)

if __name__ == "__main__":
    main()

# vim: set sts=4 sw=4 fdm=marker:
# vim: fdt=v\:folddashes\ .\ "\ "\ .\ substitute(getline(v\:foldstart+1),\ '^\\s\\+\\|#\\s*\\|\:',\ '',\ 'g'):
