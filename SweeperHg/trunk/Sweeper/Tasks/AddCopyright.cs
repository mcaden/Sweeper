namespace Sweeper
{
    using System;
    using System.Diagnostics;
    using System.Text;
    using System.Text.RegularExpressions;
    using EnvDTE;

    /// <summary>
    /// Formats the document based on visual studio formatter settings.
    /// </summary>
    public class AddCopyright : StyleTaskBase
    {
        /// <summary>
        /// Initializes a new instance of the AddCopyright class.
        /// </summary>
        public AddCopyright()
        {
            TaskName = "Add Copyright";
            TaskDescription = "Ensures that the copyright is at the top of the file.";
            Debug.WriteLine("Task: " + TaskName + " created.");
        }

        /// <summary>
        /// Gets or sets a value indicating whether the task is enabled or not.
        /// </summary>
        public override bool IsEnabled
        {
            get
            {
                return Properties.Settings.Default.AddCopyrightEnabled;
            }

            set
            {
                Properties.Settings.Default.AddCopyrightEnabled = value;
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
                    TextDocument projectItemDoc = (TextDocument)ideWindow.Document.Object("TextDocument");
                    EditPoint startPoint = projectItemDoc.StartPoint.CreateEditPoint();
                    EditPoint blockPoint = startPoint.CreateEditPoint();
                    TextRanges trs = null;
                    bool found = blockPoint.FindPattern("<copyright", (int)vsFindOptions.vsFindOptionsMatchCase, ref blockPoint, ref trs);
                    if (!found || blockPoint.Line > 2)
                    {
                        Debug.WriteLine("Finding assembly info...");
                        ProjectItem assemblyInfo = projectItem.DTE.Solution.FindProjectItem("AssemblyInfo.cs");
                        TextDocument objTextDoc = (TextDocument)ideWindow.Document.Object("TextDocument");

                        assemblyInfo.Open(EnvDTE.Constants.vsViewKindTextView).Activate();
                        TextDocument assemblyDoc = (TextDocument)assemblyInfo.DTE.ActiveWindow.Document.Object("TextDocument");

                        string company = RetrieveAssemblyInfoValue(assemblyDoc, "AssemblyCompany");
                        string copyright = RetrieveAssemblyInfoValue(assemblyDoc, "AssemblyCopyright");
                        string fileName = projectItem.Name;

                        Debug.WriteLine("Adding Copyright to file: " + projectItem.Name);
                        AddCopyrightToFile(projectItemDoc, fileName, company, copyright);
                    }
                    else
                    {
                        Debug.WriteLine("Copyright detected, skipping copyright header.");
                    }
                }
                catch (Exception exc)
                {
                    Debug.WriteLine(exc.ToString());
                    Debug.WriteLine("Format failed, skipping");
                }
            }
        }

        /// <summary>
        /// Retrieves a field from the project's assembly information.
        /// </summary>
        /// <param name="doc">The document object of the AssemblyInfo.cs file.</param>
        /// <param name="assemblyInfoKey">The AssemblyInfo key to get the data for.</param>
        /// <returns>The value of the field.</returns>
        private string RetrieveAssemblyInfoValue(TextDocument doc, string assemblyInfoKey)
        {
            string assemblyInfoValue = string.Empty;

            EditPoint startPoint = doc.StartPoint.CreateEditPoint();
            EditPoint blockPoint = startPoint.CreateEditPoint();
            TextRanges trs = null;

            if (startPoint.FindPattern(string.Format(@"\[assembly\: {0}\("".*""\)\]", assemblyInfoKey), (int)vsFindOptions.vsFindOptionsRegularExpression, ref blockPoint, ref trs))
            {
                startPoint.StartOfLine();
                blockPoint.EndOfLine();
                assemblyInfoValue = startPoint.GetText(blockPoint).Trim();
            }

            return Regex.Match(assemblyInfoValue, @""".*""").Value;
        }

        /// <summary>
        /// Adds the standard copyright blurb to the top of a given file.
        /// </summary>
        /// <param name="doc">The document object to add the copyright to.</param>
        /// <param name="fileName">The name of the file</param>
        /// <param name="company">The name of the company</param>
        /// <param name="copyright">The copyright string</param>
        private void AddCopyrightToFile(TextDocument doc, string fileName, string company, string copyright)
        {
            EditPoint startPoint = doc.StartPoint.CreateEditPoint();

            StringBuilder copyrightHeader = new StringBuilder();
            copyrightHeader.AppendLine("// -----------------------------------------------------------------------");
            copyrightHeader.AppendLine(string.Format(@"// <copyright file=""{0}"" company={1}>", fileName, company));
            copyrightHeader.AppendLine(string.Format(@"// {0}", copyright));
            copyrightHeader.AppendLine("// </copyright>");
            copyrightHeader.AppendLine("// -----------------------------------------------------------------------");
            startPoint.Insert(copyrightHeader.ToString());
        }
    }
}