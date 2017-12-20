namespace Sweeper
{
    using System;
    using System.Diagnostics;

    /// <summary>
    /// Formats the document based on visual studio formatter settings.
    /// </summary>
    public class RunStyleCopRescanAll : SolutionStyleTaskBase
    {
        /// <summary>
        /// Initializes a new instance of the RunStyleCopRescanAll class.
        /// </summary>
        public RunStyleCopRescanAll()
        {
            TaskName = "Run StyleCop (Rescan All)";
            TaskDescription = "Runs StyleCop (Rescan All) on the entire solution";
            Debug.WriteLine("Task: " + TaskName + " created.");
        }

        /// <summary>
        /// Gets or sets a value indicating whether the task is enabled or not.
        /// </summary>
        public override bool IsEnabled
        {
            get
            {
                return Properties.Settings.Default.RunStyleCopRescanAllEnabled;
            }

            set
            {
                Properties.Settings.Default.RunStyleCopRescanAllEnabled = value;
            }
        }

        /// <summary>
        /// Performs the style task.
        /// </summary>
        /// <param name="ideWindow">The IDE window.</param>
        protected override void DoWork(EnvDTE.Window ideWindow)
        {
            try
            {
                Debug.WriteLine("Running StyleCop (Rescan All) on Solution");
                ideWindow.SetFocus();
                ideWindow.DTE.ExecuteCommand("Tools.RunStyleCopRescanAll", string.Empty);
            }
            catch (Exception exc)
            {
                Debug.WriteLine(exc.ToString());
                Debug.WriteLine("StyleCop (Rescan All) failed, skipping");
            }
        }
    }
}