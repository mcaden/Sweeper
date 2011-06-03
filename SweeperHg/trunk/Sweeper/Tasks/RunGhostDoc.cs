namespace Sweeper
{
    using System;
    using System.Diagnostics;
    using EnvDTE;

    /// <summary>
    /// Formats the document based on visual studio formatter settings.
    /// </summary>
    public class RunGhostDoc : StyleTaskBase
    {
        /// <summary>
        /// Instance of the window this task is running off of.
        /// </summary>
        private EnvDTE.Window window = null;

        /// <summary>
        /// Initializes a new instance of the RunGhostDoc class.
        /// </summary>
        public RunGhostDoc()
        {
            TaskName = "Run GhostDoc";
            TaskDescription = "Runs GhostDoc on a given document if it's installed.";
            Debug.WriteLine("Task: " + TaskName + " created.");
        }

        /// <summary>
        /// Gets or sets a value indicating whether the task is enabled or not.
        /// </summary>
        public override bool IsEnabled
        {
            get
            {
                return Properties.Settings.Default.RunGhostDocEnabled;
            }

            set
            {
                Properties.Settings.Default.RunGhostDocEnabled = value;
            }
        }

        /// <summary>
        /// Performs the style task.
        /// </summary>
        /// <param name="projectItem">The project Item</param>
        /// <param name="ideWindow">The IDE window.</param>
        protected override void DoWork(ProjectItem projectItem, EnvDTE.Window ideWindow)
        {
            if (projectItem.Name.EndsWith(".cs") && ideWindow != null)
            {
                try
                {
                    Debug.WriteLine("Running GhostDoc Document: " + projectItem.Name);
                    window = ideWindow;
                    TextDocument objTextDoc = (TextDocument)ideWindow.Document.Object("TextDocument");
                    DocumentElements(objTextDoc, projectItem.FileCodeModel.CodeElements);
                }
                catch (Exception exc)
                {
                    Debug.WriteLine(exc.ToString());
                    Debug.WriteLine("Format failed, skipping");
                }
            }
        }

        /// <summary>
        /// Documents a specific element.
        /// </summary>
        /// <param name="document">The TextDocument that contains the code.</param>
        /// <param name="element">The element to document.</param>
        private void DocumentElement(EnvDTE.TextDocument document, EnvDTE.CodeElement element)
        {
            EnvDTE.EditPoint startPoint;
            document.Selection.GotoLine(element.StartPoint.Line, true);
            startPoint = document.CreateEditPoint(document.Selection.ActivePoint);
            document.Selection.MoveToPoint(startPoint, true);
            window.SetFocus();
            window.DTE.ExecuteCommand("Tools.Submain.GhostDoc.DocumentThis", string.Empty);
        }

        /// <summary>
        /// Documents child CodeElements within a document.
        /// </summary>
        /// <param name="document">The TextDocument that contains the code.</param>
        /// <param name="elements">The elements to document.</param>
        private void DocumentElements(EnvDTE.TextDocument document, EnvDTE.CodeElements elements)
        {
            try
            {
                foreach (CodeElement element in elements)
                {
                    switch (element.Kind)
                    {
                        case vsCMElement.vsCMElementFunction:
                            EnvDTE.CodeFunction func;
                            func = (EnvDTE.CodeFunction)element;
                            if (func.Access != vsCMAccess.vsCMAccessProject)
                            {
                                DocumentElement(document, element);
                            }

                            DocumentElements(document, func.Children);
                            break;
                        case vsCMElement.vsCMElementProperty:
                            EnvDTE.CodeProperty prop;
                            prop = (EnvDTE.CodeProperty)element;
                            if (prop.Access != vsCMAccess.vsCMAccessProject)
                            {
                                DocumentElement(document, element);
                            }

                            DocumentElements(document, prop.Children);
                            break;
                        case vsCMElement.vsCMElementEvent:
                            EnvDTE80.CodeEvent evt;
                            evt = (EnvDTE80.CodeEvent)element;
                            if (evt.Access != vsCMAccess.vsCMAccessProject)
                            {
                                DocumentElement(document, element);
                            }

                            DocumentElements(document, evt.Children);
                            break;
                        case vsCMElement.vsCMElementClass:
                            EnvDTE.CodeClass cls;
                            cls = (EnvDTE.CodeClass)element;
                            if (cls.Access != vsCMAccess.vsCMAccessProject)
                            {
                                DocumentElement(document, element);
                            }

                            DocumentElements(document, cls.Children);
                            break;
                        case vsCMElement.vsCMElementStruct:
                            EnvDTE.CodeStruct strct;
                            strct = (EnvDTE.CodeStruct)element;
                            if (strct.Access != vsCMAccess.vsCMAccessProject)
                            {
                                DocumentElement(document, element);
                            }

                            DocumentElements(document, strct.Children);
                            break;
                        case vsCMElement.vsCMElementDelegate:
                            EnvDTE.CodeDelegate dlg;
                            dlg = (EnvDTE.CodeDelegate)element;
                            if (dlg.Access != vsCMAccess.vsCMAccessProject)
                            {
                                DocumentElement(document, element);
                            }

                            DocumentElements(document, dlg.Children);
                            break;
                        case vsCMElement.vsCMElementEnum:
                            EnvDTE.CodeEnum enm;
                            enm = (EnvDTE.CodeEnum)element;
                            if (enm.Access != vsCMAccess.vsCMAccessProject)
                            {
                                DocumentElement(document, element);
                            }

                            DocumentElements(document, enm.Children);
                            break;
                        case vsCMElement.vsCMElementVariable:
                            EnvDTE.CodeVariable var;

                            var = (EnvDTE.CodeVariable)element;
                            if (var.Access != vsCMAccess.vsCMAccessProject)
                            {
                                DocumentElement(document, element);
                            }

                            break;
                        case vsCMElement.vsCMElementNamespace:
                            EnvDTE.CodeNamespace nmspc;
                            nmspc = (EnvDTE.CodeNamespace)element;
                            DocumentElements(document, nmspc.Children);
                            break;
                        case vsCMElement.vsCMElementInterface:
                            EnvDTE.CodeInterface inter;
                            inter = (EnvDTE.CodeInterface)element;
                            if (inter.Access != vsCMAccess.vsCMAccessProject)
                            {
                                DocumentElement(document, element);
                            }

                            DocumentElements(document, inter.Children);
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
            }
        }
    }
}