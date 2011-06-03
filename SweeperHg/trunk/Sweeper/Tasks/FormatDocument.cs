namespace Sweeper
{
    using System;
    using System.Diagnostics;
    using EnvDTE;

    /// <summary>
    /// Formats the document based on visual studio formatter settings.
    /// </summary>
    public class FormatDocument : StyleTaskBase
    {
        /// <summary>
        /// Initializes a new instance of the FormatDocument class.
        /// </summary>
        public FormatDocument()
        {
            TaskName = "Format Document";
            TaskDescription = "Formats the document based on visual studio formatter settings.";
            Debug.WriteLine("Task: " + TaskName + " created.");
        }

        /// <summary>
        /// Gets or sets a value indicating whether the task is enabled or not.
        /// </summary>
        public override bool IsEnabled
        {
            get
            {
                return Properties.Settings.Default.FormatDocumentEnabled;
            }

            set
            {
                Properties.Settings.Default.FormatDocumentEnabled = value;
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
                    Debug.WriteLine("Formatting Document: " + projectItem.Name);
                    ideWindow.SetFocus();
                    ideWindow.DTE.ExecuteCommand("Edit.FormatDocument", string.Empty);
                }
                catch (Exception exc)
                {
                    Debug.WriteLine(exc.ToString());
                    Debug.WriteLine("Format failed, skipping");
                }
            }
        }
    }
}