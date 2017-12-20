namespace Sweeper
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using EnvDTE;

    /// <summary>
    /// Moves using statements to within a namespace.
    /// </summary>
    public class MoveUsings : StyleTaskBase
    {
        /// <summary>
        /// Initializes a new instance of the MoveUsings class.
        /// </summary>
        public MoveUsings()
        {
            TaskName = "Move Usings";
            TaskDescription = "Moves Using Statements to within the namespace";
            Debug.WriteLine("Task: " + TaskName + " created.");
        }

        /// <summary>
        /// Gets or sets a value indicating whether the task is enabled or not.
        /// </summary>
        public override bool IsEnabled
        {
            get
            {
                return Properties.Settings.Default.MoveUsingsEnabled;
            }

            set
            {
                Properties.Settings.Default.MoveUsingsEnabled = value;
            }
        }

        /// <summary>
        /// Performs the style task.
        /// </summary>
        /// <param name="projectItem">The project Item</param>
        /// <param name="ideWindow">The IDE window.</param>
        protected override void DoWork(ProjectItem projectItem, EnvDTE.Window ideWindow)
        {
            if (projectItem.Name.EndsWith(".cs") && !projectItem.Name.EndsWith("Designer.cs"))
            {
                Debug.WriteLine("Moving Usings on: " + projectItem.Name);
                TextDocument objTextDoc = (TextDocument)ideWindow.Document.Object("TextDocument");
                EditPoint startPoint = objTextDoc.StartPoint.CreateEditPoint();
                string backupText = startPoint.GetText(objTextDoc.EndPoint);
                try
                {
                    List<EditPoint> namespaceInsertionPoints = GetInsertionPoints(projectItem.FileCodeModel);

                    if (namespaceInsertionPoints != null && namespaceInsertionPoints.Count > 0)
                    {
                        MoveUsingStatements(projectItem.FileCodeModel, namespaceInsertionPoints);
                    }
                }
                catch (Exception exc)
                {
                    Debug.WriteLine(exc.ToString());
                    Debug.WriteLine("Attempting to Revert...");
                    startPoint.Delete(objTextDoc.EndPoint);
                    startPoint.Insert(backupText);
                    Debug.WriteLine("Reverted.");
                }
            }
        }

        /// <summary>
        /// Gets a list of insertion points.
        /// </summary>
        /// <param name="fileCodeModel">The code model to search.</param>
        /// <returns>A list of insertion points.</returns>
        private List<EditPoint> GetInsertionPoints(FileCodeModel fileCodeModel)
        {
            List<EditPoint> namespacePoints = new List<EditPoint>();

            for (int i = 1; i <= fileCodeModel.CodeElements.Count; i++)
            {
                if (fileCodeModel.CodeElements.Item(i).Kind == vsCMElement.vsCMElementNamespace)
                {
                    EditPoint namespacePoint = fileCodeModel.CodeElements.Item(i).StartPoint.CreateEditPoint();
                    TextRanges trs = null;

                    if (namespacePoint.FindPattern("{", (int)vsFindOptions.vsFindOptionsMatchCase, ref namespacePoint, ref trs))
                    {
                        namespacePoints.Add(namespacePoint);
                    }
                }
            }

            return namespacePoints;
        }

        /// <summary>
        /// Moves using statements to a given set of insertion points.
        /// </summary>
        /// <param name="fileCodeModel">The code model to search.</param>
        /// <param name="insertionPoints">The insertion points to use.</param>
        private void MoveUsingStatements(FileCodeModel fileCodeModel, List<EditPoint> insertionPoints)
        {
            List<string> usingStatements = new List<string>();
            for (int i = 1; i <= fileCodeModel.CodeElements.Count; i++)
            {
                CodeElement element = fileCodeModel.CodeElements.Item(i);
                if (element.Kind == vsCMElement.vsCMElementImportStmt)
                {
                    EditPoint usingStart = element.StartPoint.CreateEditPoint();
                    string statement = usingStart.GetText(element.EndPoint);
                    usingStatements.Add(statement);
                    usingStart.Delete(element.EndPoint);
                    usingStart.DeleteWhitespace(vsWhitespaceOptions.vsWhitespaceOptionsVertical);
                    i--;
                }
            }

            for (int i = 0; i < insertionPoints.Count; i++)
            {
                EditPoint origin = insertionPoints[i].CreateEditPoint();
                if (usingStatements.Count > 0)
                {
                    origin.Insert("\r\n");
                    for (int j = 0; j < usingStatements.Count; j++)
                    {
                        origin.Insert(usingStatements[j] + "\r\n");
                    }
                }

                origin.DeleteWhitespace(vsWhitespaceOptions.vsWhitespaceOptionsVertical);

                if (usingStatements.Count > 0)
                {
                    origin.Insert("\r\n");
                }
            }
        }
    }
}
