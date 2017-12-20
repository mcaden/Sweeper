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
                Debug.WriteLine("Correcting Block Spacing: " + projectItem.Name);
                TextDocument objTextDoc = (TextDocument)ideWindow.Document.Object("TextDocument");
                EditPoint startPoint = objTextDoc.StartPoint.CreateEditPoint();
                foreach (CodeElement element in projectItem.FileCodeModel.CodeElements)
                {
                    FormatBlockSpacing(element);
                }
            }
        }

        /// <summary>
        /// Removes whitespace immediately inside a block
        /// </summary>
        /// <param name="element">The current code element</param>
        private void RemoveInternalBlockPadding(CodeElement element)
        {
            EditPoint start = element.GetStartPoint(vsCMPart.vsCMPartBody).CreateEditPoint();
            EditPoint end = element.GetEndPoint(vsCMPart.vsCMPartBody).CreateEditPoint();
            end.DeleteWhitespace(vsWhitespaceOptions.vsWhitespaceOptionsVertical);
            end.CharLeft(1);

            start.DeleteWhitespace(vsWhitespaceOptions.vsWhitespaceOptionsVertical);
            end.DeleteWhitespace(vsWhitespaceOptions.vsWhitespaceOptionsVertical);
        }

        /// <summary>
        /// Checks to see if there should be additional blank lines immediately after starting a block
        /// </summary>
        /// <param name="element">The current element to check</param>
        private void CheckBlockStart(CodeElement element)
        {
            EditPoint start = element.GetStartPoint(vsCMPart.vsCMPartBody).CreateEditPoint();
            EditPoint end = element.GetEndPoint(vsCMPart.vsCMPartBody).CreateEditPoint();

            EditPoint beginStart = start.CreateEditPoint();
            beginStart.StartOfLine();

            string beginningStartText = beginStart.GetText(start).Trim();
            if (beginningStartText != string.Empty)
            {
                EditPoint endStart = start.CreateEditPoint();
                endStart.EndOfLine();
                string restofStartText = start.GetText(endStart).Trim();
                if (!restofStartText.StartsWith("get;"))
                {
                    start.Insert(Environment.NewLine);
                    start.SmartFormat(endStart);
                }
            }
        }

        /// <summary>
        /// Checks to see if there should be additional blank lines after the end of a block.
        /// </summary>
        /// <param name="element">The current element to check</param>
        private void CheckBlockEnd(CodeElement element)
        {
            EditPoint endBlock = element.GetEndPoint(vsCMPart.vsCMPartWholeWithAttributes).CreateEditPoint();
            EditPoint startBlock = element.GetStartPoint(vsCMPart.vsCMPartWholeWithAttributes).CreateEditPoint();
            string original = startBlock.GetText(endBlock);

            EditPoint endOfEnd = endBlock.CreateEditPoint();
            endOfEnd.EndOfLine();
            string endOfBlockLine = endBlock.GetText(endOfEnd).Trim();
            if (element.Kind == vsCMElement.vsCMElementAttribute || element.Kind == vsCMElement.vsCMElementOther)
            {
                if (endOfBlockLine != "]" && endOfBlockLine != ")]" && !endOfBlockLine.StartsWith(","))
                {
                    endOfEnd = endBlock.CreateEditPoint();
                    endOfEnd.EndOfLine();
                    endOfBlockLine = endBlock.GetText(endOfEnd).Trim();
                    if (endOfBlockLine != string.Empty)
                    {
                        endBlock.Insert(Environment.NewLine);
                        endBlock.SmartFormat(endOfEnd);
                    }
                }
            }
            else if (endOfBlockLine != string.Empty)
            {
                endBlock.Insert(Environment.NewLine);
                endBlock.SmartFormat(endOfEnd);
            }

            if (element.Kind != vsCMElement.vsCMElementImportStmt &&
                element.Kind != vsCMElement.vsCMElementAttribute &&
                element.Kind != vsCMElement.vsCMElementOther)
            {
                endOfEnd.LineDown(1);
                endOfEnd.EndOfLine();
                string lineAfterBlock = endBlock.GetText(endOfEnd).Trim();

                if (lineAfterBlock != string.Empty && !lineAfterBlock.StartsWith("else") && !lineAfterBlock.StartsWith("}"))
                {
                    endBlock.Insert(Environment.NewLine);
                    endBlock.SmartFormat(endOfEnd);
                }
            }
        }

        /// <summary>
        /// Removes Blank lines after the opening of a block, or right before the closing of a block.
        /// </summary>
        /// <param name="element">The current code element</param>
        private void FormatBlockSpacing(CodeElement element)
        {
            if (element.Kind != vsCMElement.vsCMElementImportStmt &&
                element.Kind != vsCMElement.vsCMElementVariable &&
                element.Kind != vsCMElement.vsCMElementEvent &&
                element.Kind != vsCMElement.vsCMElementParameter &&
                element.Kind != vsCMElement.vsCMElementAttribute &&
                element.Kind != vsCMElement.vsCMElementOther)
            {
                RemoveInternalBlockPadding(element);
                CheckBlockStart(element);
            }

            if (element.Kind != vsCMElement.vsCMElementParameter)
            {
                CheckBlockEnd(element);

                foreach (CodeElement childElement in element.Children)
                {
                    if (childElement.Kind == vsCMElement.vsCMElementFunction)
                    {
                        EditPoint headerStart = childElement.GetStartPoint(vsCMPart.vsCMPartHeader).CreateEditPoint();
                        EditPoint headerEnd = headerStart.CreateEditPoint();
                        headerEnd.EndOfLine();
                        string header = headerStart.GetText(headerEnd).Trim();
                        if (header.StartsWith("partial"))
                        {
                            continue;
                        }
                    }

                    FormatBlockSpacing(childElement);
                }
            }
        }
    }
}