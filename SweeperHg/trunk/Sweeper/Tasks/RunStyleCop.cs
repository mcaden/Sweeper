namespace Sweeper
{
    using System;
    using System.Diagnostics;
    using EnvDTE;

    /// <summary>
    /// Formats the document based on visual studio formatter settings.
    /// </summary>
    public class RunStyleCop : StyleTaskBase
    {
        /// <summary>
        /// Initializes a new instance of the RunStyleCop class.
        /// </summary>
        public RunStyleCop()
        {
            TaskName = "Run StyleCop";
            TaskDescription = "Runs StyleCop on a single ProjectItem";
            Debug.WriteLine("Task: " + TaskName + " created.");
        }

        /// <summary>
        /// Gets or sets a value indicating whether the task is enabled or not.
        /// </summary>
        public override bool IsEnabled
        {
            get
            {
                return Properties.Settings.Default.RunStyleCopEnabled;
            }

            set
            {
                Properties.Settings.Default.RunStyleCopEnabled = value;
            }
        }

        /// <summary>
        /// Performs the style task.
        /// </summary>
        /// <param name="projectItem">The project Item</param>
        /// <param name="ideWindow">The IDE window.</param>
        protected override void DoWork(ProjectItem projectItem, EnvDTE.Window ideWindow)
        {
            if (projectItem.Name.EndsWith(".cs"))
            {
                try
                {
                    Debug.WriteLine("Running StyleCop: " + projectItem.Name);
                    ideWindow.SetFocus();
                    ideWindow.DTE.ExecuteCommand("Tools.RunStyleCop", string.Empty);
                }
                catch (Exception exc)
                {
                    Debug.WriteLine(exc.ToString());
                    Debug.WriteLine("StyleCop failed, skipping");
                }
            }
        }
    }
}