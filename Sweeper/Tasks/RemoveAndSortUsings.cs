namespace Sweeper
{
    using System;
    using System.Diagnostics;
    using EnvDTE;

    /// <summary>
    /// Calls the Edit.RemoveAndSort macro on a given ProjectItem
    /// </summary>
    public class RemoveAndSortUsings : StyleTaskBase
    {
        /// <summary>
        /// Functions used to detect errors.  The macro doesn't run if there are errors.
        /// </summary>
        private ErrorDetectionFunction errorDetectionFunction = null;

        /// <summary>
        /// Initializes a new instance of the RemoveAndSortUsings class.
        /// </summary>
        /// <param name="errorDetector">ErrorDetectionFunction delegate used to detect errors.  Pass in null to skip this.</param>
        public RemoveAndSortUsings(ErrorDetectionFunction errorDetector)
        {
            TaskName = "Remove and Sort Usings";
            TaskDescription = "Removes unnecessary Using statements and sorts the usings with the system namespace first and all other namespaces after it alphabetically.";
            Debug.WriteLine("Task: " + TaskName + " created.");
            errorDetectionFunction = errorDetector;
        }

        /// <summary>
        /// Error detection to determine if the macro should be skipped.
        /// </summary>
        /// <returns>True if errors have been detected.</returns>
        public delegate bool ErrorDetectionFunction();

        /// <summary>
        /// Gets or sets a value indicating whether the task is enabled or not.
        /// </summary>
        public override bool IsEnabled
        {
            get
            {
                return Properties.Settings.Default.RemoveAndSortUsingsEnabled;
            }

            set
            {
                Properties.Settings.Default.RemoveAndSortUsingsEnabled = value;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the task is disabled for a single iteration.
        /// </summary>
        public bool IsTemporarilyDisabled { get; set; }

        /// <summary>
        /// Resets temporary disabling that would affect the remainder of a pass, but not the next pass.
        /// </summary>
        public override void Reset()
        {
            IsTemporarilyDisabled = false;
        }

        /// <summary>
        /// Performs the style task.
        /// </summary>
        /// <param name="projectItem">The project Item</param>
        /// <param name="ideWindow">The IDE window.</param>
        protected override void DoWork(ProjectItem projectItem, EnvDTE.Window ideWindow)
        {
            if (projectItem.Name.EndsWith(".cs") && !IsTemporarilyDisabled)
            {
                try
                {
                    if (errorDetectionFunction != null)
                    {
                        if (errorDetectionFunction())
                        {
                            IsTemporarilyDisabled = true;
                            return;
                        }
                    }

                    Debug.WriteLine("Removing and Sorting Usings on: " + projectItem.Name);

                    Stopwatch macroWatch = new Stopwatch();
                    macroWatch.Start();

                    ideWindow.SetFocus();
                    ideWindow.DTE.ExecuteCommand("Edit.RemoveAndSort", string.Empty);

                    macroWatch.Stop();
                    if (macroWatch.ElapsedMilliseconds > 1000)
                    {
                        if (System.Windows.Forms.MessageBox.Show("Remove and Sort usings seems to be having difficulties, do you want to skip it this iteration?", "Parole", System.Windows.Forms.MessageBoxButtons.YesNo) == System.Windows.Forms.DialogResult.Yes)
                        {
                            IsTemporarilyDisabled = true;
                        }
                    }
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