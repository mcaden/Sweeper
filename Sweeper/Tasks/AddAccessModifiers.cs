namespace Sweeper
{
    using System.Diagnostics;
    using System.Reflection;
    using EnvDTE;
    using EnvDTE80;

    /// <summary>
    /// Adds access modifiers to CodeElements that are lacking it except those that shouldn't have them.
    /// </summary>
    public class AddAccessModifiers : StyleTaskBase
    {
        /// <summary>
        /// Initializes a new instance of the AddAccessModifiers class
        /// </summary>
        public AddAccessModifiers()
        {
            TaskName = "Add Missing Access Modifiers";
            TaskDescription = "Adds missing access modifiers.";
            Debug.WriteLine("Task: " + TaskName + " created.");
        }

        /// <summary>
        /// Gets or sets a value indicating whether the task is enabled or not.
        /// </summary>
        public override bool IsEnabled
        {
            get
            {
                return Properties.Settings.Default.AddAccessModifiersEnabled;
            }

            set
            {
                Properties.Settings.Default.AddAccessModifiersEnabled = value;
            }
        }

        /// <summary>
        /// Performs the style task.
        /// </summary>
        /// <param name="projectItem">The project Item</param>
        /// <param name="ideWindow">The IDE window.</param>
        protected override void DoWork(ProjectItem projectItem, EnvDTE.Window ideWindow)
        {
            Debug.WriteLine("Adding missing access modifiers: " + projectItem.Name);
            FileCodeModel fileCodeModel = projectItem.FileCodeModel;
            for (int i = 1; i <= fileCodeModel.CodeElements.Count; i++)
            {
                if (fileCodeModel.CodeElements.Item(i).Kind == vsCMElement.vsCMElementNamespace)
                {
                    CodeElement element = fileCodeModel.CodeElements.Item(i);

                    AddMissingAccessModifiers(element as CodeElement);
                }
            }
        }

        /// <summary>
        /// Adds missing access modifiers
        /// </summary>
        /// <param name="codeElement">The CodeElement to add missing access modifiers too. Includes children.</param>
        private void AddMissingAccessModifiers(CodeElement codeElement)
        {
            if (codeElement.Kind != vsCMElement.vsCMElementInterface && CodeElementBlockTypes.ContainsKey(codeElement.Kind))
            {
                for (int i = 1; i <= codeElement.Children.Count; i++)
                {
                    CodeElement element = codeElement.Children.Item(i) as CodeElement;

                    if (element.Kind != vsCMElement.vsCMElementImportStmt && element.Kind != vsCMElement.vsCMElementInterface && CodeElementBlockTypes.ContainsKey(codeElement.Kind))
                    {
                        // Get the element's access through reflection rather than a massive switch.
                        vsCMAccess access = (vsCMAccess)CodeElementTypes[element.Kind].InvokeMember("Access", BindingFlags.GetProperty, null, element, null);

                        if (element.Kind == vsCMElement.vsCMElementClass || element.Kind == vsCMElement.vsCMElementStruct)
                        {
                            AddMissingAccessModifiers(element);
                        }

                        EditPoint start;
                        EditPoint end;
                        string declaration;

                        if (element.Kind == vsCMElement.vsCMElementFunction)
                        {
                            // method, constructor, or finalizer
                            CodeFunction2 function = element as CodeFunction2;

                            // Finalizers don't have access modifiers, neither do static constructors
                            if (function.FunctionKind == vsCMFunction.vsCMFunctionDestructor || (function.FunctionKind == vsCMFunction.vsCMFunctionConstructor && function.IsShared))
                            {
                                continue;
                            }
                        }

                        if (element.Kind == vsCMElement.vsCMElementProperty || element.Kind == vsCMElement.vsCMElementVariable || element.Kind == vsCMElement.vsCMElementEvent)
                        {
                            CodeElements attributes = (CodeElements)CodeElementTypes[element.Kind].InvokeMember("Attributes", BindingFlags.GetProperty, null, element, null);

                            start = attributes.Count > 0 ? element.GetEndPoint(vsCMPart.vsCMPartAttributesWithDelimiter).CreateEditPoint() : element.StartPoint.CreateEditPoint();
                        }
                        else
                        {
                            start = element.GetStartPoint(vsCMPart.vsCMPartHeader).CreateEditPoint();
                        }

                        end = start.CreateEditPoint();
                        end.EndOfLine();
                        declaration = start.GetText(end);
                        if (!declaration.StartsWith(CodeAccessKeywords[access]) && !declaration.StartsWith("partial"))
                        {
                            object[] args = new object[1];
                            args.SetValue(access, 0);
                            CodeElementTypes[element.Kind].InvokeMember("Access", BindingFlags.SetProperty, null, element, args);
                        }

                        declaration = start.GetText(end);
                    }
                }
            }
        }
    }
}
