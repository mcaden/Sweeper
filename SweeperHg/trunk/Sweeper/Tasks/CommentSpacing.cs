namespace Sweeper
{
    using System;
    using System.Diagnostics;
    using EnvDTE;

    /// <summary>
    /// Formats the document based on visual studio formatter settings.
    /// </summary>
    public class CommentSpacing : StyleTaskBase
    {
        /// <summary>
        /// Initializes a new instance of the CommentSpacing class.
        /// </summary>
        public CommentSpacing()
        {
            TaskName = "Format Comment Spacing";
            TaskDescription = "Formats the spacing around comments.";
            Debug.WriteLine("Task: " + TaskName + " created.");
        }

        /// <summary>
        /// Gets or sets a value indicating whether the task is enabled or not.
        /// </summary>
        public override bool IsEnabled
        {
            get
            {
                return Properties.Settings.Default.CommentSpacing;
            }

            set
            {
                Properties.Settings.Default.CommentSpacing = value;
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
                Debug.WriteLine("Formatting Spacing Around Comments: " + projectItem.Name);
                try
                {
                    TextDocument objTextDoc = (TextDocument)ideWindow.Document.Object("TextDocument");
                    EditPoint objEditPoint = objTextDoc.CreateEditPoint(objTextDoc.StartPoint);
                    EditPoint commentPoint = objEditPoint.CreateEditPoint();
                    TextRanges trs = null;

                    while (objEditPoint.FindPattern("//", (int)vsFindOptions.vsFindOptionsMatchCase, ref commentPoint, ref trs))
                    {
                        bool previousBlank = false;
                        EditPoint previousCheckPoint = objEditPoint.CreateEditPoint();
                        previousCheckPoint.LineUp(1);
                        if (previousCheckPoint.GetText(objEditPoint).Trim() == string.Empty)
                        {
                            previousBlank = true;
                        }

                        commentPoint.CharRight(1);
                        string comment = objEditPoint.GetText(commentPoint);
                        while (!comment.EndsWith(" ") && !commentPoint.AtEndOfLine)
                        {
                            if (comment.EndsWith("/"))
                            {
                                commentPoint.CharRight(1);
                            }
                            else
                            {
                                commentPoint.CharLeft(1);
                                commentPoint.Insert(" ");
                            }

                            comment = objEditPoint.GetText(commentPoint);
                        }

                        commentPoint.CharRight(1);
                        comment = objEditPoint.GetText(commentPoint);
                        if (comment.EndsWith("  "))
                        {
                            commentPoint.CharLeft(1);
                            commentPoint.DeleteWhitespace(vsWhitespaceOptions.vsWhitespaceOptionsHorizontal);
                            commentPoint.Insert(" ");
                        }

                        if (commentPoint.Line > objEditPoint.Line)
                        {
                            commentPoint.LineUp(1);
                            commentPoint.EndOfLine();
                        }

                        if (commentPoint.AtEndOfLine)
                        {
                            objEditPoint.Delete(commentPoint);
                        }
                        else
                        {
                            EditPoint endComment = commentPoint.CreateEditPoint();
                            endComment.EndOfLine();
                            if (commentPoint.GetText(endComment).Trim() == string.Empty)
                            {
                                objEditPoint.Delete(endComment);
                            }
                            else
                            {
                                objEditPoint.LineDown(1);
                                previousBlank = false;
                            }
                        }

                        objEditPoint.StartOfLine();
                        commentPoint = objEditPoint.CreateEditPoint();
                        commentPoint.EndOfLine();
                        if (objEditPoint.GetText(commentPoint).Trim() == string.Empty)
                        {
                            objEditPoint.DeleteWhitespace(vsWhitespaceOptions.vsWhitespaceOptionsVertical);
                            if (previousBlank)
                            {
                                objEditPoint.Insert("\r\n");
                            }
                        }
                    }
                }
                catch (Exception exc)
                {
                    Debug.WriteLine(exc.ToString());
                    Debug.WriteLine("Formatting Spacing Around Comments failed, skipping");
                }
            }
        }
    }
}