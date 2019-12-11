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
        public static readonly Guid CommandSet = new Guid("7fd09a49-455e-442f-921e-c2c4d5e12997");

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
                var menuCommandID = new CommandID(CommandSet, CommandId);
                var menuItem = new MenuCommand(this.MenuItemCallback, menuCommandID);
                commandService.AddCommand(menuItem);
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
            try
            {
                Project solutionFolder = null;

                DirectoryInfo folder = GetFolder();
                if (folder == null)
                    return;

                Solution2 solution = (Solution2)dte.Solution;
                var projects = solution.Projects;

                foreach (Project item in projects)
                {
                    if (item.Name == folder.Name)
                    {
                        solutionFolder = item;
                        break;
                    }
                }

                if (solutionFolder == null)
                {
                    solutionFolder = solution.AddSolutionFolder(folder.Name);
                }

                dte.StatusBar.Text = $"Creating Solution Folders for {folder.Name}";

                IncludeFiles(folder, solutionFolder);

                dte.StatusBar.Text = $"Created Solution Folders for {folder.Name}";
            }
            catch (Exception ex)
            {
                dte.StatusBar.Text = $"Error while Creating Solution Folders: {ex.Message}";
            }
        }

        private static void IncludeFiles(DirectoryInfo folder, Project project)
        {
            foreach (var item in folder.GetFileSystemInfos())
            {
                if (item is FileInfo)
                {
                    project.ProjectItems.AddFromFile(item.FullName);
                }
                else if (item is DirectoryInfo)
                {
                    // Skip git repo folder since it's too big and not necessary.
                    if (item.Name == ".git")
                        continue;

                    var solutionFolder = (SolutionFolder)project.Object;
                    var newSolutionFolder = solutionFolder.AddSolutionFolder(item.Name);
                    IncludeFiles((DirectoryInfo)item, newSolutionFolder);
                }
            }
        }

        private DirectoryInfo GetFolder()
        {
            Solution2 sol2 = (Solution2)dte.Solution;
            var solutionPath = Path.GetDirectoryName(sol2.FullName);
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.ShowNewFolderButton = false;
                dialog.Description = "Select a folder. The folder you pick will be created as a solution folder and containing files will be added to it.";
                dialog.SelectedPath = solutionPath;

                var r = dialog.ShowDialog();

                if (r == DialogResult.OK)
                    return new DirectoryInfo(dialog.SelectedPath);
            }

            return null;
        }
    }
}
