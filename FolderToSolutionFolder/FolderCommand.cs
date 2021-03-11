using System;
using System.ComponentModel.Design;
using System.Globalization;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System.IO;
using Microsoft.Win32;
using System.Windows.Forms;
using EnvDTE;
using EnvDTE80;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio;
using Microsoft.WindowsAPICodePack.Dialogs;



namespace CeciliaSharp.FolderToSolutionFolder
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class FolderCommand
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSetSolution = new Guid("7fd09a49-455e-442f-921e-c2c4d5e12997");
        public static readonly Guid CommandSetProject = new Guid("7fd09a49-455e-442f-921e-c2c4d5e12998");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly Package package;
        private static Package serviceProvider;
        private static DTE2 dte;

        /// <summary>
        /// Initializes a new instance of the <see cref="FolderCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        private FolderCommand(Package package)
        {
            this.package = package ?? throw new ArgumentNullException("package");

            if (this.ServiceProvider.GetService(typeof(IMenuCommandService)) is OleMenuCommandService commandService)
            {
                {
                    var menuCommandID = new CommandID(CommandSetSolution, CommandId);
                    var menuItem = new MenuCommand(this.MenuItemCallback, menuCommandID);
                    commandService.AddCommand(menuItem);
                }
                {
                    var menuCommandID = new CommandID(CommandSetProject, CommandId);
                    var menuItem = new MenuCommand(this.MenuItemCallback, menuCommandID);
                    commandService.AddCommand(menuItem);
                }
            }
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static FolderCommand Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private IServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static void Initialize(Package package)
        {
            serviceProvider = package;
            dte = ((IServiceProvider)serviceProvider).GetService(typeof(SDTE)) as DTE2;
            Instance = new FolderCommand(package);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void MenuItemCallback(object sender, EventArgs e)
        {
            string message = string.Format(CultureInfo.CurrentCulture, "Inside {0}.MenuItemCallback()", this.GetType().FullName);

            // figure out if a project or project-folder was clicked
            ProjectItems items = null;
            try
            {
                IntPtr hierarchyPointer, selectionContainerPointer;
                Object selectedObject = null;
                IVsMultiItemSelect multiItemSelect;
                uint projectItemId;

                IVsMonitorSelection monitorSelection =
                        (IVsMonitorSelection)Package.GetGlobalService(
                        typeof(SVsShellMonitorSelection));

                monitorSelection.GetCurrentSelection(out hierarchyPointer,
                                                     out projectItemId,
                                                     out multiItemSelect,
                                                     out selectionContainerPointer);

                IVsHierarchy selectedHierarchy = Marshal.GetTypedObjectForIUnknown(
                                                     hierarchyPointer,
                                                     typeof(IVsHierarchy)) as IVsHierarchy;

                if (selectedHierarchy != null)
                {
                    ErrorHandler.ThrowOnFailure(selectedHierarchy.GetProperty(
                                                      projectItemId,
                                                      (int)__VSHPROPID.VSHPROPID_ExtObject,
                                                      out selectedObject));
                }

                dynamic dyn = selectedObject;
                items = (ProjectItems)(dyn.ProjectItems);
            }
            catch(Exception)
            { 
                // assuming solution was clicked
            }


            try
            {

                DirectoryInfo folder = GetFolder();
                if (folder == null)
                    return;

                bool isSolutionFolder = false;
                const string folderKindVirtual  = @"{6BB5F8F0-4483-11D3-8BCF-00C04F8EC28C}";
                
                string folderKind = folderKindVirtual;
                if (items != null)
                {
                    items = items.AddFolder(folder.Name, folderKind).ProjectItems;
                }
                else
                {
                    isSolutionFolder = true;
                    Solution2 solution = (Solution2)dte.Solution;
                    items = solution.AddSolutionFolder(folder.Name).ProjectItems;
                }

                dte.StatusBar.Text = $"Creating Folders for {folder.Name}";

                IncludeFiles(folder, items, isSolutionFolder, folderKind);

                dte.StatusBar.Text = $"Created Folders for {folder.Name}";
            }
            catch (Exception ex)
            {
                dte.StatusBar.Text = $"Error while Creating Folders: {ex.Message}";
            }
        }

        private static void IncludeFiles(DirectoryInfo folder, ProjectItems items, bool isSolutionFolder, string folderKind)
        {
            foreach (var item in folder.GetFileSystemInfos())
            {
                if (item is FileInfo)
                {
                    try
                    {
                        items.AddFromFile(item.FullName);
                    }
                    catch(Exception)
                    {
                        throw new Exception($"Cannot add file: {item.FullName}");
                    }
                }
                else if (item is DirectoryInfo info)
                {
                    if (isSolutionFolder)
                    {
                        var solutionFolder = (SolutionFolder)((Project)items.Parent).Object;
                        var newSolutionFolder = solutionFolder.AddSolutionFolder(item.Name).ProjectItems;
                        IncludeFiles(info, newSolutionFolder, isSolutionFolder, folderKind);
                    }
                    else
                    {
                        var newItems = items.AddFolder(item.Name, folderKind).ProjectItems;
                        IncludeFiles(info, newItems, isSolutionFolder, folderKind);
                    }
                }
            }
        }

        private DirectoryInfo GetFolder()
        {
            Solution2 sol2 = (Solution2)dte.Solution;
            var solutionPath = Path.GetDirectoryName(sol2.FullName);

            using (var dialog = new CommonOpenFileDialog
            {
                IsFolderPicker = true,
                Title = "Select a folder. The folder you pick will be created as a solution folder and containing files will be added to it.",
                InitialDirectory = solutionPath

            })

            //using (var dialog = new FolderBrowserDialog())
            {
                var r = dialog.ShowDialog();

                if (r == CommonFileDialogResult.Ok)
                    return new DirectoryInfo(dialog.FileName);
            }

            return null;
        }
    }
}
