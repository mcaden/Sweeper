namespace Sweeper
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Reflection;
    using System.Resources;
    using EnvDTE;
    using EnvDTE80;
    using Extensibility;
    using Microsoft.VisualStudio.CommandBars;

    /// <summary>The object for implementing an Add-in.</summary>
    /// <seealso class='IDTExtensibility2' />
    public class Connect : IDTExtensibility2, IDTCommandTarget
    {
        /// <summary>
        /// The application Project
        /// </summary>
        private DTE2 applicationObject;

        /// <summary>
        /// The addin instance.
        /// </summary>
        private AddIn addInInstance;

        /// <summary>
        /// The Sweeper outlook pane.
        /// </summary>
        private OutputWindowPane sweeperPane = null;

        /// <summary>
        /// A collection of the style tasks to run.
        /// </summary>
        private List<IProjectItemStyleTask> commonTasks = new List<IProjectItemStyleTask>();

        /// <summary>
        /// A collection of the style tasks to run after running sweeper on a single object.
        /// </summary>
        private List<IProjectItemStyleTask> postSingleTasks = new List<IProjectItemStyleTask>();

        /// <summary>
        /// A collection of the style tasks to run after running sweeper on the entire solution.
        /// </summary>
        private List<ISolutionStyleTask> postSolutionTasks = new List<ISolutionStyleTask>();

        /// <summary>
        /// Initializes a new instance of the Connect class as the Add-in object.
        /// </summary>
        public Connect()
        {
            commonTasks.Add(new MoveUsings());
            commonTasks.Add(new RemoveAndSortUsings(IsCompilerErrorInErrorList));
            commonTasks.Add(new SortElementsWithinClass());
            commonTasks.Add(new FormatDocument());
            commonTasks.Add(new RemoveMultipleBlankLines());
            commonTasks.Add(new BlockSpacing());
            commonTasks.Add(new AddAccessModifiers());
            commonTasks.Add(new RunGhostDoc());
            commonTasks.Add(new CommentSpacing());
            commonTasks.Add(new AddCopyright());
            postSingleTasks.Add(new RunStyleCop());
            postSolutionTasks.Add(new RunStyleCopRescanAll());
        }

        /// <summary>
        /// Gets a valid Sweeper Output Pane.
        /// </summary>
        protected OutputWindowPane SweeperPane
        {
            get
            {
                OutputWindow outputWindow = applicationObject.ToolWindows.OutputWindow;
                for (int i = 1; i <= outputWindow.OutputWindowPanes.Count; i++)
                {
                    if (outputWindow.OutputWindowPanes.Item(i).Name == "Sweeper")
                    {
                        sweeperPane = outputWindow.OutputWindowPanes.Item(i);
                        break;
                    }
                }

                if (sweeperPane == null)
                {
                    sweeperPane = outputWindow.OutputWindowPanes.Add("Sweeper");
                }

                return sweeperPane;
            }
        }

        /// <summary>Implements the OnConnection method of the IDTExtensibility2 interface. Receives notification that the Add-in is being loaded.</summary>
        /// <param name='application'>Root object of the host application.</param>
        /// <param name='connectMode'>Describes how the Add-in is being loaded.</param>
        /// <param name='addInInst'>Object representing this Add-in.</param>
        /// <param name="custom">Custom array.</param>
        /// <seealso class='IDTExtensibility2' />
        public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            applicationObject = (DTE2)application;
            addInInstance = (AddIn)addInInst;
            if (connectMode == ext_ConnectMode.ext_cm_UISetup)
            {
                object[] contextGUIDS = new object[] { };
                Commands2 commands = (Commands2)applicationObject.Commands;
                string toolsMenuName;

                try
                {
                    // If you would like to move the command to a different menu, change the word "Tools" to the 
                    //  English version of the menu. This code will take the culture, append on the name of the menu
                    //  then add the command to that menu. You can find a list of all the top-level menus in the file
                    //  CommandBar.resx.
                    string resourceName;
                    ResourceManager resourceManager = new ResourceManager("Sweeper.CommandBar", Assembly.GetExecutingAssembly());
                    CultureInfo cultureInfo = new CultureInfo(applicationObject.LocaleID);

                    if (cultureInfo.TwoLetterISOLanguageName == "zh")
                    {
                        System.Globalization.CultureInfo parentCultureInfo = cultureInfo.Parent;
                        resourceName = string.Concat(parentCultureInfo.Name, "Tools");
                    }
                    else
                    {
                        resourceName = string.Concat(cultureInfo.TwoLetterISOLanguageName, "Tools");
                    }

                    toolsMenuName = resourceManager.GetString(resourceName);
                }
                catch
                {
                    // We tried to find a localized version of the word Tools, but one was not found.
                    //  Default to the en-US word, which may work for the current culture.
                    toolsMenuName = "Tools";
                }

                // Place the command on the tools menu.
                // Find the MenuBar command bar, which is the top-level command bar holding all the main menu items:
                Microsoft.VisualStudio.CommandBars.CommandBar menuBarCommandBar = ((Microsoft.VisualStudio.CommandBars.CommandBars)applicationObject.CommandBars)["MenuBar"];
                Microsoft.VisualStudio.CommandBars.CommandBar contextCommands = ((Microsoft.VisualStudio.CommandBars.CommandBars)applicationObject.CommandBars)["Code Window"];
                Microsoft.VisualStudio.CommandBars.CommandBar solutionContextCommands = ((Microsoft.VisualStudio.CommandBars.CommandBars)applicationObject.CommandBars)["Solution"];
                Microsoft.VisualStudio.CommandBars.CommandBar itemContextCommands = ((Microsoft.VisualStudio.CommandBars.CommandBars)applicationObject.CommandBars)["Item"];

                // Find the Tools command bar on the MenuBar command bar:
                CommandBarControl toolsControl = menuBarCommandBar.Controls[toolsMenuName];
                CommandBarPopup toolsPopup = (CommandBarPopup)toolsControl;

                /* // If we want to add a popup to the context menu, use this:
                // CommandBarPopup contextPopup = (CommandBarPopup)contextCommands.Controls.Add(MsoControlType.msoControlPopup, System.Type.Missing, System.Type.Missing, 1, System.Type.Missing);
                // contextPopup.Caption = "Sweeper"; */

                // This try/catch block can be duplicated if you wish to add multiple commands to be handled by your Add-in,
                // just make sure you also update the QueryStatus/Exec method to include the new command names.
                try
                {
                    // Add a command to the Commands collection:
                    Command command = commands.AddNamedCommand2(addInInstance, "SweeperAll", "Sweeper - All", "Calls in the Sweeper on all project items.", true, 19, ref contextGUIDS, (int)vsCommandStatus.vsCommandStatusSupported + (int)vsCommandStatus.vsCommandStatusEnabled, (int)vsCommandStyle.vsCommandStylePictAndText, vsCommandControlType.vsCommandControlTypeButton);

                    // Add a control for the command to the tools menu:
                    if (command != null && toolsPopup != null)
                    {
                        command.AddControl(toolsPopup.CommandBar, 1);
                    }

                    if (command != null && solutionContextCommands != null)
                    {
                        command.AddControl(solutionContextCommands, 1);
                    }
                }
                catch (System.ArgumentException)
                {
                    //// If we are here, then the exception is probably because a command with that name
                    ////  already exists. If so there is no need to recreate the command and we can 
                    ////  safely ignore the exception.
                }

                try
                {
                    // Add a command to the Commands collection:
                    Command command = commands.AddNamedCommand2(addInInstance, "Sweeper", "Sweeper", "Calls in the Sweeper on this project item.", true, 5, ref contextGUIDS, (int)vsCommandStatus.vsCommandStatusSupported + (int)vsCommandStatus.vsCommandStatusEnabled, (int)vsCommandStyle.vsCommandStylePictAndText, vsCommandControlType.vsCommandControlTypeButton);

                    if (command != null)
                    {
                        // Add a control for the command to the tools menu:
                        if (toolsPopup != null)
                        {
                            command.AddControl(toolsPopup.CommandBar, 1);
                        }

                        if (contextCommands != null)
                        {
                            command.AddControl(contextCommands, 1);
                        }

                        if (itemContextCommands != null)
                        {
                            command.AddControl(itemContextCommands, 1);
                        }
                    }
                }
                catch (System.ArgumentException)
                {
                    //// If we are here, then the exception is probably because a command with that name
                    ////  already exists. If so there is no need to recreate the command and we can 
                    ////  safely ignore the exception.
                }

                try
                {
                    // Add a command to the Commands collection:
                    Command command = commands.AddNamedCommand2(addInInstance, "SweeperOptions", "Sweeper Options...", "Sets Sweeper Options.", true, 6, ref contextGUIDS, (int)vsCommandStatus.vsCommandStatusSupported + (int)vsCommandStatus.vsCommandStatusEnabled, (int)vsCommandStyle.vsCommandStylePictAndText, vsCommandControlType.vsCommandControlTypeButton);

                    // Add a control for the command to the tools menu:
                    if ((command != null) && (toolsPopup != null))
                    {
                        command.AddControl(toolsPopup.CommandBar, 1);
                    }
                }
                catch (System.ArgumentException)
                {
                    //// If we are here, then the exception is probably because a command with that name
                    ////  already exists. If so there is no need to recreate the command and we can 
                    ////  safely ignore the exception.
                }
            }
        }

        /// <summary>Implements the OnDisconnection method of the IDTExtensibility2 interface. Receives notification that the Add-in is being unloaded.</summary>
        /// <param name='disconnectMode'>Describes how the Add-in is being unloaded.</param>
        /// <param name='custom'>Array of parameters that are host application specific.</param>
        /// <seealso class='IDTExtensibility2' />
        public void OnDisconnection(ext_DisconnectMode disconnectMode, ref Array custom)
        {
        }

        /// <summary>Implements the OnAddInsUpdate method of the IDTExtensibility2 interface. Receives notification when the collection of Add-ins has changed.</summary>
        /// <param name='custom'>Array of parameters that are host application specific.</param>
        /// <seealso class='IDTExtensibility2'/>
        public void OnAddInsUpdate(ref Array custom)
        {
        }

        /// <summary>Implements the OnStartupComplete method of the IDTExtensibility2 interface. Receives notification that the host application has completed loading.</summary>
        /// <param name='custom'>Array of parameters that are host application specific.</param>
        /// <seealso class='IDTExtensibility2' />
        public void OnStartupComplete(ref Array custom)
        {
        }

        /// <summary>Implements the OnBeginShutdown method of the IDTExtensibility2 interface. Receives notification that the host application is being unloaded.</summary>
        /// <param name='custom'>Array of parameters that are host application specific.</param>
        /// <seealso class='IDTExtensibility2' />
        public void OnBeginShutdown(ref Array custom)
        {
            Properties.Settings.Default.Save();
        }

        /// <summary>Implements the QueryStatus method of the IDTCommandTarget interface. This is called when the command's availability is updated</summary>
        /// <param name='commandName'>The name of the command to determine state for.</param>
        /// <param name='neededText'>Text that is needed for the command.</param>
        /// <param name='status'>The state of the command in the user interface.</param>
        /// <param name='commandText'>Text requested by the neededText parameter.</param>
        /// <seealso class='Exec' />
        public void QueryStatus(string commandName, vsCommandStatusTextWanted neededText, ref vsCommandStatus status, ref object commandText)
        {
            if (neededText == vsCommandStatusTextWanted.vsCommandStatusTextWantedNone)
            {
                if (commandName == "Sweeper.Connect.SweeperAll")
                {
                    status = (vsCommandStatus)vsCommandStatus.vsCommandStatusSupported | vsCommandStatus.vsCommandStatusEnabled;
                    return;
                }

                if (commandName == "Sweeper.Connect.Sweeper")
                {
                    status = (vsCommandStatus)vsCommandStatus.vsCommandStatusSupported | vsCommandStatus.vsCommandStatusEnabled;
                    return;
                }

                if (commandName == "Sweeper.Connect.SweeperOptions")
                {
                    status = (vsCommandStatus)vsCommandStatus.vsCommandStatusSupported | vsCommandStatus.vsCommandStatusEnabled;
                    return;
                }
            }
        }

        /// <summary>Implements the Exec method of the IDTCommandTarget interface. This is called when the command is invoked.</summary>
        /// <param name='commandName'>The name of the command to execute.</param>
        /// <param name='executeOption'>Describes how the command should be run.</param>
        /// <param name='varIn'>Parameters passed from the caller to the command handler.</param>
        /// <param name='varOut'>Parameters passed from the command handler to the caller.</param>
        /// <param name='handled'>Informs the caller if the command was handled or not.</param>
        /// <seealso class='Exec' />
        public void Exec(string commandName, vsCommandExecOption executeOption, ref object varIn, ref object varOut, ref bool handled)
        {
            ProjectItem projectItem = applicationObject.ActiveWindow.ProjectItem;
            handled = false;
            if (executeOption == vsCommandExecOption.vsCommandExecOptionDoDefault)
            {
                if (commandName == "Sweeper.Connect.SweeperAll")
                {
                    ClearOutputWindow();
                    WriteLineToOutputWindow("-= Starting Sweeper =-");

                    Projects theProjects = applicationObject.Solution.Projects;

                    if (theProjects.Count > 0)
                    {
                        foreach (Project theProject in theProjects)
                        {
                            if (theProject.ProjectItems != null && theProject.ProjectItems.Count > 0)
                            {
                                foreach (ProjectItem p in theProject.ProjectItems)
                                {
                                    RunProjectItemStyleTasks(commonTasks, p);
                                }
                            }
                            else
                            {
                                WriteLineToOutputWindow("! Skipping Project: " + theProject.Name + " - No Items in project. !");
                            }
                        }

                        RunSolutionStyleTasks(postSolutionTasks);
                        ResetTasks(commonTasks);
                        ResetTasks(postSolutionTasks);
                    }
                    else
                    {
                        WriteLineToOutputWindow("! Not Running - No Projects in solution !");
                    }

                    WriteLineToOutputWindow("-= Sweeper Complete =-");

                    if (projectItem != null)
                    {
                        Window theWindow = projectItem.Open(Constants.vsViewKindCode);
                        theWindow.Activate();
                    }

                    handled = true;
                    return;
                }
                else if (commandName == "Sweeper.Connect.Sweeper")
                {
                    CommandBars commandBars = (CommandBars)applicationObject.CommandBars;

                    ClearOutputWindow();
                    WriteLineToOutputWindow("-= Starting Sweeper =-");

                    Document activeDoc = applicationObject.ActiveDocument;
                    
                    if (activeDoc != null)
                    {
                        ProjectItem p = activeDoc.ProjectItem;
                        if (p != null)
                        {
                            RunProjectItemStyleTasks(commonTasks, p);
                            ResetTasks(commonTasks);

                            RunProjectItemStyleTasks(postSingleTasks, p);
                            ResetTasks(postSingleTasks);
                        }
                        else
                        {
                            WriteLineToOutputWindow("! Not Running - Active Document isn't a valid project item !");
                        }
                    }
                    else
                    {
                        WriteLineToOutputWindow("! Not Running - No Active Document !");
                    }

                    WriteLineToOutputWindow("-= Sweeper Complete =-");

                    Window theWindow = activeDoc.ProjectItem.Open(Constants.vsViewKindCode);
                    theWindow.Activate();
                    handled = true;
                    return;
                }
                else if (commandName == "Sweeper.Connect.SweeperOptions")
                {
                    List<object> taskOptions = new List<object>();
                    for (int i = 0; i < commonTasks.Count; i++)
                    {
                        taskOptions.Add(commonTasks[i]);
                    }

                    for (int i = 0; i < postSingleTasks.Count; i++)
                    {
                        taskOptions.Add(postSingleTasks[i]);
                    }

                    for (int i = 0; i < postSolutionTasks.Count; i++)
                    {
                        taskOptions.Add(postSolutionTasks[i]);
                    }

                    SweeperUI.OptionWindow optionWindow = new SweeperUI.OptionWindow(taskOptions);
                    bool dialogResult = optionWindow.ShowDialog().Value;
                    System.Diagnostics.Debug.WriteLine(dialogResult.ToString());

                    Window theWindow = projectItem.Open(Constants.vsViewKindCode);
                    theWindow.Activate();
                    handled = true;
                    return;
                }
            }
        }

        /// <summary>
        /// Resets all temporary settings for a set of tasks.
        /// </summary>
        /// <typeparam name="T">An IStyleTask derivative</typeparam>
        /// <param name="styleTasks">A list of styleTasks to reset</param>
        public void ResetTasks<T>(List<T> styleTasks) where T : IStyleTask
        {
            for (int i = 0; i < styleTasks.Count; i++)
            {
                styleTasks[i].Reset();
            }
        }

        /// <summary>
        /// Runs the style tasks populated within the tasks field.
        /// </summary>
        /// <param name="styleTasks">The set of tasks to run on the project item.</param>
        /// <param name="projectItem">The project item to run the tasks on.</param>
        public void RunProjectItemStyleTasks(List<IProjectItemStyleTask> styleTasks, ProjectItem projectItem)
        {
            if (projectItem.Name.EndsWith(".cs") && !projectItem.Name.EndsWith("Designer.cs"))
            {
                bool isItemOpen = projectItem.get_IsOpen(Constants.vsViewKindCode);
                Window theWindow = projectItem.Open(Constants.vsViewKindCode);
                theWindow.Activate();

                WriteLineToOutputWindow("\r\n-- Working on File: " + projectItem.Name + "--");

                try
                {
                    for (int i = 0; i < styleTasks.Count; i++)
                    {
                        if (styleTasks[i].IsEnabled)
                        {
                            WriteLineToOutputWindow("Running: " + styleTasks[i].TaskName);
                            styleTasks[i].PerformStyleTask(projectItem, theWindow);
                        }
                        else
                        {
                            WriteLineToOutputWindow("Not Running: " + styleTasks[i].TaskName + " -- Disabled");
                        }
                    }
                }
                catch (Exception e)
                {
                    System.Diagnostics.Trace.WriteLine(e.Message.ToString());
                }

                if (!isItemOpen && projectItem.Saved)
                {
                    theWindow.Close(vsSaveChanges.vsSaveChangesPrompt);
                }
            }

            if (projectItem.ProjectItems != null)
            {
                foreach (ProjectItem p in projectItem.ProjectItems)
                {
                    RunProjectItemStyleTasks(styleTasks, p);
                }
            }
        }

        /// <summary>
        /// Runs the style tasks populated within the tasks field.
        /// </summary>
        /// <param name="styleTasks">The set of tasks to run on the project item.</param>
        public void RunSolutionStyleTasks(List<ISolutionStyleTask> styleTasks)
        {
            EnvDTE.Window theWindow = applicationObject.ActiveWindow;
            theWindow.Activate();

            WriteLineToOutputWindow("\r\n-- Post: Running Solution Tasks --");

            try
            {
                for (int i = 0; i < styleTasks.Count; i++)
                {
                    if (styleTasks[i].IsEnabled)
                    {
                        WriteLineToOutputWindow("Running: " + styleTasks[i].TaskName);
                        styleTasks[i].PerformStyleTask(theWindow);
                    }
                    else
                    {
                        WriteLineToOutputWindow("Not Running: " + styleTasks[i].TaskName + " -- Disabled");
                    }
                }
            }
            catch (Exception e)
            {
                System.Diagnostics.Trace.WriteLine(e.Message.ToString());
            }
        }

        /// <summary>
        /// Checks the error list for compiler errors.
        /// </summary>
        /// <returns>True if there's something in the list.</returns>
        private bool IsCompilerErrorInErrorList()
        {
            ErrorList myErrors;

            applicationObject.ExecuteCommand("View.ErrorList", " ");
            myErrors = applicationObject.ToolWindows.ErrorList;
            for (int i = 1; i <= myErrors.ErrorItems.Count; i++)
            {
                if (myErrors.ErrorItems.Item(i).ErrorLevel == vsBuildErrorLevel.vsBuildErrorLevelHigh)
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// Writes a given string to the Sweeper Output pane.
        /// </summary>
        /// <param name="outputString">The string to write to the pane.</param>
        private void WriteLineToOutputWindow(string outputString)
        {
            SweeperPane.OutputString(outputString + "\r\n");
            applicationObject.ExecuteCommand("View.Output", " ");
            SweeperPane.Activate();
        }

        /// <summary>
        /// Clears the Sweeper Output pane.
        /// </summary>
        private void ClearOutputWindow()
        {
            SweeperPane.Clear();
            SweeperPane.Activate();
        }
    }
}