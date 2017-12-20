namespace Sweeper
{
    using System;
    using System.Diagnostics;
    using EnvDTE;

    /// <summary>
    /// Task to remove multiple blank lines.
    /// </summary>
    public class RemoveMultipleBlankLines : StyleTaskBase
    {
        /// <summary>
        /// Initializes a new instance of the RemoveMultipleBlankLines class
        /// </summary>
        public RemoveMultipleBlankLines()
        {
            TaskName = "Remove Multiple Blank Lines";
            TaskDescription = "Consolidates multiple blank lines down into 1 single blank line.";
            Debug.WriteLine("Task: " + TaskName + " created.");
        }

        /// <summary>
        /// Gets or sets a value indicating whether the task is enabled or not.
        /// </summary>
        public override bool IsEnabled
        {
            get
            {
                return Properties.Settings.Default.RemoveMultipleBlankLinesEnabled;
            }

            set
            {
                Properties.Settings.Default.RemoveMultipleBlankLinesEnabled = value;
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
                Debug.WriteLine("Removing Unnecessary Blank Lines: " + projectItem.Name);
                try
                {
                    TextDocument objTextDoc = (TextDocument)ideWindow.Document.Object("TextDocument");
                    EditPoint objEditPoint = objTextDoc.CreateEditPoint(objTextDoc.StartPoint);
                    while (!objEditPoint.AtEndOfDocument)
                    {
                        int secondFarthestLine = objEditPoint.Line + 2;
                        if (secondFarthestLine > objTextDoc.EndPoint.Line)
                        {
                            secondFarthestLine = objEditPoint.Line + 1;
                        }

                        if (objEditPoint.GetLines(objEditPoint.Line, secondFarthestLine).Trim() == string.Empty)
                        {
                            objEditPoint.DeleteWhitespace(vsWhitespaceOptions.vsWhitespaceOptionsVertical);
                            objEditPoint.Insert("\r\n");
                        }

                        objEditPoint.LineDown(1);
                    }
                }
                catch (Exception exc)
                {
                    Debug.WriteLine(exc.ToString());
                    Debug.WriteLine("Removing Unnecessary Blank Lines failed, skipping");
                }
            }
        }
    }
}