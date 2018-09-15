using EnvDTE;
using EnvDTE80;
using LocateFolder.Utilities;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Globalization;
using System.IO;
using System.Linq;
using Task = System.Threading.Tasks.Task;

namespace LocateFolder
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class LocateFolderCommand
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int LFCommandId = 0x0100;
        public const int OBFCommandId = 0x3B9ACA01;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid LocateFolderCommandSet = new Guid("f31449ab-b753-48fe-bf3f-25c691905367");
        public static readonly Guid OpenBinFolderCommandSet = new Guid("02AB237F-F580-4278-A02B-8DA88483528E");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        private DTE2 _applicationObject { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="LocateFolderCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private LocateFolderCommand(AsyncPackage package, OleMenuCommandService commandService, DTE2 appObj)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            this._applicationObject = appObj ?? throw new ArgumentNullException(nameof(_applicationObject));

            var menuCommandID = new CommandID(LocateFolderCommandSet, LFCommandId);
            var menuItem = new MenuCommand(this.LocateFolder, menuCommandID);
            commandService.AddCommand(menuItem);

            menuCommandID = new CommandID(OpenBinFolderCommandSet, OBFCommandId);
            menuItem = new MenuCommand(this.OpenBinFolderWithFileExplorer, menuCommandID);
            commandService.AddCommand(menuItem);
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static LocateFolderCommand Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider
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
        public static async Task InitializeAsync(AsyncPackage package)
        {
            // Switch to the main thread - the call to AddCommand in LocateFolderCommand's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync((typeof(IMenuCommandService))) as OleMenuCommandService;
            DTE2 applicationObject = await package.GetServiceAsync(typeof(DTE)) as DTE2;

            Instance = new LocateFolderCommand(package, commandService, applicationObject);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void LocateFolder(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            UIHierarchy uih = _applicationObject.ToolWindows.SolutionExplorer;
            var selectedItems = ((Array)uih.SelectedItems);

            if (selectedItems != null)
            {
                LocateFile.FilesOrFolders((IEnumerable<string>)(from object t in selectedItems
                                                                where (t as UIHierarchyItem)?.Object is ProjectItem
                                                                select ((ProjectItem)((UIHierarchyItem)t).Object).FileNames[1]));
            }
        }

        private void OpenBinFolderWithFileExplorer(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            //Get the active projects within the solution.
            Array _activeProjects = (Array)_applicationObject.ActiveSolutionProjects;

            //loop through each active project
            foreach (Project _activeProject in _activeProjects)
            {
                //get the directory path based on the project file.
                string _projectPath = Path.GetDirectoryName(_activeProject.FullName);
                //get the output path based on the active configuration
                string _projectOutputPath = _activeProject.ConfigurationManager.ActiveConfiguration.Properties.Item("OutputPath").Value.ToString();
                //combine the project path and output path to get the bin path
                string _projectBinPath = Path.Combine(_projectPath, _projectOutputPath);

                //if the directory exists (already built) then open that directory
                //in windows explorer using the diagnostics.process object
                if (Directory.Exists(_projectBinPath))
                {
                    System.Diagnostics.Process.Start(_projectBinPath);
                }
                else
                {
                    //if the directory doesnt exist, open the project directory.
                    System.Diagnostics.Process.Start(_projectPath);
                }
            }
        }

    }
}