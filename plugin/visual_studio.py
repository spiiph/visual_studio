############################################################ {{{1
# Documentation
'''\
visual_studio.py - Companion file to visual_studio.vim
Version: 2.0
Author: Henrik Öhman <speeph@gmail.com>
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
# Global objects
wsh = WScriptShell()
dte = DTEWrapper(wsh)
vimext = VimExt(wsh)

############################################################ {{{1
# Exit function
# Disable vim exit processing to avoid an exception at exit
sys.exitfunc = lambda: None

############################################################ {{{1
# Logging
import logging
logging_enabled = True

if logging_enabled:
    import tempfile
    log_file = os.path.join(tempfile.gettempdir(), "visual_studio.log")
    logging.basicConfig(
            filename = log_file,
            filemode = "w",
            level = logging.DEBUG,
            format = ("%(asctime)s %(levelname)-8s %(pathname)s" +
                "(%(lineno)d)\n%(message)s\n")
            )
else:
    logging.basicConfig(level = sys.maxint)
logging.info("starting")

############################################################ {{{1
class DTEWrapper:
    '''The DTE class encapsulates DTE objects and functionality.'''

    ############################################################ {{{2
    # Initialization and properties
    def __init__(self, wsh):
        # Dict containing {pid: dte} pairs
        self.dtes = {}

        # The pid of the current DTE object
        self.current_dte = 0

        # WScript.Shell
        self.wsh = wsh

    def __get_dte():
        if self.current_dte == 0:
            self.set_current_dte()
        return self.dtes[self.current_dte]

    dte = property(__get_dte)
        

    ############################################################ {{{2
    def set_current_dte(self, pid = 0):
        '''Get the DTE object corresponding to pid. If pid is 0, get the
        current DTE object. If the current DTE object is None, get the first
        DTE object in the self.dtes dict.'''
        pid = int(pid)

        if pid == 0 or pid == self.current_dte:
            if self.dtes[self.current_dte] is not None:
                try:
                    self.dtes[self.current_dte].Solution
                    return self.dtes[self.current_dte]
                except:
                    del self.dtes[self.current_dte]

        try:
            self.update_dtes()
            if self.dtes.has_key(pid):
                self.current_dte = pid
            else:
                self.current_dte = self.dtes.keys()[0]
            return self.dtes[self.current_dte]
        except pywintypes.com_error:
            VimExt.echomsg("Error: Cannot access DTE object. " +
                "Is Visual Studio running?")
            return None

    ############################################################ {{{2
    def update_dtes(self):
        '''Populate the self.dtes dict with {pid: dte} elements from the
        Running Object Table.'''
        self.dtes = {}
        rot = pythoncom.GetRunningObjectTable()
        rot_enum = rot.EnumRunning()
        context = pythoncom.CreateBindCtx(0)
        while 1:
            monikers = rot_enum.Next()
            if monikers is None:
                break

            display_name = monikers[0].GetDisplayName(context, None)
            if display_name.startswith("!VisualStudio.DTE"):
                logging.info("== DTE.update_dtes display_name: %s"
                        % display_name)

                try:
                    pid = int(display_name.rpartition(":")[-1])
                except ValueError:
                    pid = 0

                dte = win32com.client.Dispatch(
                        rot.GetObject(monikers[0]).QueryInterface(
                            pythoncom.IID_Dispatch))
                self.dtes[pid] = dte

    ############################################################ {{{2
    def wait_for_build(self):
        '''Wait for Visual Studio to complete the build.'''
        # Visual Studio build state flags
        vsBuildStateNotStarted = 1
        vsBuildStateInProgress = 2
        vsBuildStateDone = 3

        build = self.dte.Solution.SolutionBuild
        try:
            while build.BuildState == vsBuildStateInProgress:
                time.sleep(0.1)
        except:
            pass

    ############################################################ {{{2
    def activate(self):
        '''Activate Visual Studio.'''
        if self.dte is None:
            return

        try:
            self.dte.MainWindow.Activate()
            wsh().AppActivate(self.dte.MainWindow.Caption)
        except pywintypes.com_error:
            VimExt.echomsg("Error: Cannot access WScript Shell object.")

    ############################################################ {{{2
    def set_autoload(self):
        '''Activate the Autoload option in Visual Studio.'''
        if self.dte is None:
            return

        try:
            properties = self.dte.Properties("Environment", "Documents")
            enabled = []
            for item in ["DetectFileChangesOutsideIDE",
                    "AutoloadExternalChanges"]:
                if not properties.Item(item).Value:
                    p.Item(item).Value = 1
                    enabled.append(item)
        except pywintypes.com_error:
            pass
        VimExt.echo("Enabled %s in VisualStudio" % " and ".join(enabled))

    ############################################################ {{{2
    def has_csharp_projects(self):
        '''Check if a solution has C# projects.'''
        if self.dte is None:
            return

        try:
            if self.dte.CSharpProjects.Count:
                return 1
        except pywintypes.com_error, AttributeError:
            pass
        return 0

    ############################################################ {{{2
    def output(self, output_file, caption):
        '''Fetch the output from a command and write it to a file.'''
        logging.info("== DTE.output %s" % vars())

        if self.dte is None:
            return

        try:
            window = self.dte.Windows.Item(caption)
        except pywintypes.com_error:
            VimExt.echomsg("Error: window not active (%s)" % caption)
            return

        if caption == "Output":
            sel = window.Object.OutputWindowPanes.Item("Build").\
                TextDocument.Selection
        else:
            sel = window.Selection
        sel.SelectAll()
        f = file(output_file, "w")
        f.write(sel.Text)
        f.close()
        sel.Collapse()

    ############################################################ {{{2
    def compile_file(self, output_file):
        '''Compile the current file.'''
        logging.info("== DTE.compile_file %s" % vars())

        if self.dte is None:
            return

        try:
            self.dte.ExecuteCommand("Build.Compile")
        except Exception, e:
            logging.exception(e)
            VimExt.exception(e, sys.exc_traceback)
            VimExt.vim_activate()
            return

        # Wait for build to complete
        self.wait_for_build()
        self.output(output_file, "Output")
        #VimExt.vim_status("Compile file complete")
        VimExt.vim_activate()

    ############################################################ {{{2
    def build_project(self, output_file, project_name = None):
        '''Build a project by name or build the startup project.'''
        logging.info("== DTE.build_project %s" % vars())

        if self.dte is None:
            return

        if has_csharp_projects():
            self.dte.Documents.CloseAll()
        self.set_autoload()
        self.activate()
        self.output_activate()

        try:
            solution = self.dte.Solution
            config = solution.ActiveConfiguration.Name
            if project_name is None:
                project_name = solution.Properties("StartupProject").Value
            project = [x for x in solution.Projects
                    if x.Name == project_name][0]

            logging.info(("== DTE.build_project configuration_name: " +
                "%s project_unique_name: %s") % (config, project.UniqueName))

            solution.SolutionBuild.BuildProject(config, project.UniqueName, 1)
        except Exception, e:
            logging.exception(e)
            VimExt.exception(e, sys.exc_traceback)
            VimExt.activate()
            return

        # Wait for build to complete
        self.wait_for_build()
        self.output(output_file, "Output")
        #VimExt.status("Build project complete")
        VimExt.activate()

    ############################################################ {{{2
    def build_solution(self, output_file):
        '''Build the solution in the current DTE.'''
        logging.info("== DTE.build_solution %s" % vars())

        if self.dte is None:
            return

        if has_csharp_projects():
            self.dte.Documents.CloseAll()
        self.set_autoload()
        self.activate()
        self.output_activate()

        try:
            self.dte.Solution.SolutionBuild.Build(1)
        except Exception, e:
            logging.exception(e)
            VimExt.exception(e, sys.exc_traceback)
            VimExt.activate()
            return

        # Wait for build to complete
        self.wait_for_build()
        self.output(output_file, "Output")
        #VimExt.status("Build solution complete")
        VimExt.activate()

    ############################################################ {{{2
    def set_startup_project(self, project_name):
        '''Set the startup project in Visual Studio.'''
        logging.info("== DTE.set_startup_project %s" % vars())

        if self.dte is None:
            return

        try:
            self.dte.Solution.Properties("StartupProject").Value =
                project_name
        except Exception, e:
            logging.exception(e)
            VimExt.exception(e, sys.exc_traceback)
            return
        #VimExt.status("Startup project set to %s" % project_name)

    ############################################################ {{{2
    def get_file(self):
        '''Get the current file from Visual Studio.'''
        logging.info("== DTE.get_file %s" % vars())

        if self.dte is None:
            return

        doc = dte.ActiveDocument
        if doc is None
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
        logging.info("== DTE.put_file %s" % vars())

        logging.info("== DTE.put_file abspath %s" %
                (os.path.abspath(filename)))

        if self.dte is None:
            return

        self.set_autoload()
        item_op = self.dte.ItemOperations.OpenFile(
                os.path.abspath(filename))
        sel = dte.ActiveDocument.Selection
        sel.MoveToLineAndOffset(line, col)
        self.activate()

    ############################################################ {{{2
    def update_instance_list(self):
        '''Update Vim's list of solutions.'''
        logging.info("== DTE.update_instance_list %s" % vars())
        instances = [[key, self.dtes[key].Solution.FullName]
                for key in self.dtes.keys()]
        VimExt.command("let s:visual_studio_solutions = %s" % instances)

    ############################################################ {{{2
    def update_project_list(self):
        '''Update Vim's list of projects.'''
        logging.info("== DTE.update_project_list %s" % vars())

        if self.dte is None:
            return

        startup_project_name =
            self.dte.Solution.Properties("StartupProject").Value
        startup_project_index = -1
        index = 0
        projects = []
        for project in sorted(self.dte.Solution.Projects,
                cmp = lambda x,y: cmp(x.Name, y.Name)):
            if project.Name == startup_project_name:
                startup_project_index = index
            projects.append(self.get_project_tree(project))
            index += 1
        VimExt.command("let s:visual_studio_projects = %s" % projects)
        VimExt.command("let s:visual_studio_project_index = %s" %
                startup_project_index)

    ############################################################ {{{3
    def get_project_tree(self, project):
        '''Returns a tree (nested lists) of projects and files in projects.
        The first item is the project or project item name. The second item
        contains a list of children or a filename.'''
        def com_property(object, attr, default = None):
            try:
                return getattr(object, attr, default)
            except:
                return default

        name = com_property(project, "Name")
        if name is None:
            return []

        name = str(name)
        properties = com_property(project, "Properties")
        if properties is not None:
            try:
                full_path = str(properties["FullPath"])
            except:
                full_path = None
            # Item is a leaf; return name and full path
            if full_path and os.path.isfile(full_path):
                return [name, full_path]

        # Item is a sub project; recurse deeper
        sub_project = com_property(project, "SubProject")
        if sub_project:
            return [name, self.get_project_tree(sub_project)]

        # Item is a project items container; recurse deeper for each project
        # item
        project_items = com_property(project, "ProjectItems")
        if project_items:
            children = [self.get_project_tree(x) for x in project_items]
            return [name, children]

        return [name, []]

    ############################################################ {{{2
    def update_project_files_list(self, project_name = None):
        '''Update Vim's list of files for the named project or the startup
        project.'''
        logging.info("== DTE.get_project_files %s" % vars())

        if self.dte is None:
            return

        files = []
        try:
            solution = self.dte.Solution
            if not project_name:
                project_name = solution.Properties("StartupProject").Value
            project = [x for x in solution.Projects
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
            if i.Properties.Item("Kind").Value == "VCFile":
                files.append(i.Properties.Item("FullPath").Value)
            else:
                files += self.get_project_items_files(i.ProjectItems)
        return files

    ############################################################ {{{2
    def get_task_list(self, output_file):
        logging.info("== DTE.task_list %s" % vars())

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
            except:
                filename = "<no-filename>"

            try:
                line = item.Line
            except:
                line = "<no-line>"

            try:
                description = item.Description
            except:
                description = "<no-description>"

            f.write("%s(%s) : %s" % (filename, line, description))
        f.close()

############################################################ {{{1
class WScriptShell:
    def __init__(self):
        self.__wsh = win32com.client.Dispatch("WScript.Shell")

    def __call__(self, *args, **kw):
        return __wsh

############################################################ {{{1
class VimExt:
    '''Vim extension class for DTEWrapper.'''

    __pid = None
    __has_python = None

    ############################################################ {{{2
    def set_pid(pid):
        '''Set the PID of Vim when issued from the command line.'''
        VimExt.__pid = pid

    ############################################################ {{{2
    def has_python():
        '''Check if Vim was compiled with the +python feature.'''
        if VimExt.__has_python is None:
            if globals().get("vim") is None:
                VimExt.__has_python = 0
            else:
                VimExt.__has_python = int(vim.eval("has('python')"))
        return VimExt.__has_python

    ############################################################ {{{2
    def activate():
        '''Activate Vim.'''
        if VimExt.__pid is None:
            VimExt.__pid = os.getpid()
        wsh().AppActivate(VimExt.__pid)

    ############################################################ {{{2
    def command(command_list):
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

    ############################################################ {{{2
    def exception(e, trace):
        '''Print the output of an exception in Vim.'''
        if isinstance(e, pywintypes.com_error):
            try:
                msg = e[2][2]
            except:
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

    ############################################################ {{{2
    def VimExt.escape(s):
        '''Escape a string for :echo, :echomsg and :echoerr.'''
        return str(s).replace("'", "''")

    ############################################################ {{{2
    def VimExt.echo(msg):
        ''':echo a message.'''
        VimExt.command("echo '%s'" % VimExt.escape(msg))

    ############################################################ {{{2
    def VimExt.echomsg(msg):
        ''':echomsg a message.'''
        VimExt.command("echomsg '%s'" % VimExt.escape(msg))

    ############################################################ {{{2
    def VimExt.echoerr(msg):
        ''':echoerr a message.'''
        VimExt.command("echoerr '%s'" % VimExt.escape(msg))


def dte_execute(name, *args):
    if hasattr(dte, name):
        VimExt.echoerr("Error: no such function %s in %s" %
                (name, prog))
    else:
        function = getattr(dte, name)
        function(args)

############################################################ {{{2
def main():
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
        VimExt.exception(e, sys.exc_traceback)

if __name__ == "__main__":
    main()

# vim: set sts=4 sw=4 fdm=marker:
# vim: fdt=v\:folddashes\ .\ substitute(getline(v\:foldstart+1),\ '[#\:]',\ '',\ 'g'):
