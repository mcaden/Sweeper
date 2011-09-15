namespace Sweeper
{
    using System;
    using System.Collections.Generic;
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
        /// Keywords used for code element access
        /// </summary>
        private Dictionary<vsCMAccess, string> codeAccessKeywords = new Dictionary<vsCMAccess, string>() 
        { 
                { vsCMAccess.vsCMAccessPublic, "public" }, 
                { vsCMAccess.vsCMAccessProtected, "protected" },
                { vsCMAccess.vsCMAccessPrivate, "private" },
                { vsCMAccess.vsCMAccessProject, "internal" }
        };

        /// <summary>
        /// Keywords used for code element access
        /// </summary>
        private Dictionary<vsCMElement, Type> codeElementTypes = new Dictionary<vsCMElement, Type>() 
        { 
                { vsCMElement.vsCMElementClass, typeof(CodeClass) }, 
                { vsCMElement.vsCMElementDelegate, typeof(CodeDelegate) },
                { vsCMElement.vsCMElementEnum, typeof(CodeEnum) },
                { vsCMElement.vsCMElementEvent, typeof(CodeEvent) },
                { vsCMElement.vsCMElementFunction, typeof(CodeFunction) },
                { vsCMElement.vsCMElementProperty, typeof(CodeProperty) },
                { vsCMElement.vsCMElementStruct, typeof(CodeStruct) },
                { vsCMElement.vsCMElementVariable, typeof(CodeVariable) }
        };

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
            if (codeElement.Kind != vsCMElement.vsCMElementInterface)
            {
                for (int i = 1; i <= codeElement.Children.Count; i++)
                {
                    CodeElement element = codeElement.Children.Item(i) as CodeElement;

                    if (element.Kind != vsCMElement.vsCMElementImportStmt && element.Kind != vsCMElement.vsCMElementInterface)
                    {
                        // Get the element's access through reflection rather than a massive switch.
                        vsCMAccess access = (vsCMAccess)codeElementTypes[element.Kind].InvokeMember("Access", BindingFlags.GetProperty, null, element, null);

                        if (element.Kind == vsCMElement.vsCMElementClass || element.Kind == vsCMElement.vsCMElementStruct)
                        {
                            AddMissingAccessModifiers(element);
                        }

                        EditPoint start;

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
                            CodeElements attributes = (CodeElements)codeElementTypes[element.Kind].InvokeMember("Attributes", BindingFlags.GetProperty, null, element, null);

                            start = attributes.Count > 0 ? element.GetEndPoint(vsCMPart.vsCMPartAttributesWithDelimiter).CreateEditPoint() : element.StartPoint.CreateEditPoint();
                        }
                        else
                        {
                            start = element.GetStartPoint(vsCMPart.vsCMPartHeader).CreateEditPoint();
                        }

                        EditPoint end = start.CreateEditPoint();
                        end.EndOfLine();
                        string declaration = start.GetText(end);
                        if (!declaration.StartsWith(codeAccessKeywords[access]))
                        {
                            object[] args = new object[1];
                            args.SetValue(access, 0);
                            codeElementTypes[element.Kind].InvokeMember("Access", BindingFlags.SetProperty, null, element, args);
                        }

                        declaration = start.GetText(end);
                    }
                }
            }
        }
    }
}
