namespace Sweeper
{
    using System;
    using System.Diagnostics;
    using EnvDTE;

    /// <summary>
    /// Formats the document based on visual studio formatter settings.
    /// </summary>
    public class BlockSpacing : StyleTaskBase
    {
        /// <summary>
        /// Initializes a new instance of the BlockSpacing class.
        /// </summary>
        public BlockSpacing()
        {
            TaskName = "Correct Block Spacing";
            TaskDescription = "Corrects spacing at the end of a block.";
            Debug.WriteLine("Task: " + TaskName + " created.");
        }

        /// <summary>
        /// Gets or sets a value indicating whether the task is enabled or not.
        /// </summary>
        public override bool IsEnabled
        {
            get
            {
                return Properties.Settings.Default.BlockSpacingEnabled;
            }

            set
            {
                Properties.Settings.Default.BlockSpacingEnabled = value;
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
                    Debug.WriteLine("Correcting Block Spacing: " + projectItem.Name);
                    TextDocument objTextDoc = (TextDocument)ideWindow.Document.Object("TextDocument");
                    EditPoint startPoint = objTextDoc.StartPoint.CreateEditPoint();
                    RemoveOpenBlockBlankLines(objTextDoc);

                    startPoint = objTextDoc.StartPoint.CreateEditPoint();
                     RemoveCloseBlockBlankLines(objTextDoc);
                }
                catch (Exception exc)
                {
                    Debug.WriteLine(exc.ToString());
                    Debug.WriteLine("Format failed, skipping");
                }
            }
        }

        /// <summary>
        /// Removes Blank lines right after opening a block.
        /// </summary>
        /// <param name="doc">The document object to search.</param>
        private void RemoveOpenBlockBlankLines(TextDocument doc)
        {
            EditPoint startPoint = doc.StartPoint.CreateEditPoint();
            EditPoint blockPoint = startPoint.CreateEditPoint();
            TextRanges trs = null;

            while (startPoint.FindPattern("{", (int)vsFindOptions.vsFindOptionsMatchCase, ref blockPoint, ref trs))
            {
                EditPoint checkPoint = blockPoint.CreateEditPoint();
                if (checkPoint.Line <= doc.EndPoint.Line - 2)
                {
                    checkPoint.LineDown(2);
                    checkPoint.StartOfLine();
                    if (blockPoint.GetText(checkPoint).Trim() == String.Empty)
                    {
                        checkPoint.LineUp(1);
                        checkPoint.DeleteWhitespace(vsWhitespaceOptions.vsWhitespaceOptionsVertical);
                    }
                }

                startPoint = blockPoint;
            }
        }

        /// <summary>
        /// Removes Blank lines right before closing a block.
        /// </summary>
        /// <param name="doc">The document object to search.</param>
        private void RemoveCloseBlockBlankLines(TextDocument doc)
        {
            EditPoint startPoint = doc.StartPoint.CreateEditPoint();
            EditPoint blockPoint = startPoint.CreateEditPoint();
            TextRanges trs = null;

            while (startPoint.FindPattern("}", (int)vsFindOptions.vsFindOptionsMatchCase, ref blockPoint, ref trs))
            {
                EditPoint checkPoint = startPoint.CreateEditPoint();
                if (checkPoint.Line >= 3)
                {
                    checkPoint.LineUp(2);
                    checkPoint.EndOfLine();
                    if (startPoint.GetText(checkPoint).Trim() == String.Empty)
                    {
                        checkPoint.LineDown(1);
                        checkPoint.DeleteWhitespace(vsWhitespaceOptions.vsWhitespaceOptionsVertical);
                    }

                    startPoint = blockPoint;
                }
            }
        }
    }
}