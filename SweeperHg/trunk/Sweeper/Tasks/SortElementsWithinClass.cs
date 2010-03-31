namespace Sweeper
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.Windows.Forms;
    using EnvDTE;

    /// <summary>
    /// Sorts Elements according to their type and access level to fit standard.
    /// </summary>
    public class SortElementsWithinClass : StyleTaskBase
    {
        /// <summary>
        /// Initializes a new instance of the SortElementsWithinClass class.
        /// </summary>
        public SortElementsWithinClass()
        {
            TaskName = "Sort Elements Within a Class";
            TaskDescription = "Sorts Elements according to their type and access level to fit standard.";
            Debug.WriteLine("Task: " + TaskName + " created.");
            IsTemporarilyDisabled = false;
        }

        /// <summary>
        /// Class priority enums to sort elements by their type.  
        /// Built-in type enums are not sufficient for this because they span multiple types.
        /// </summary>
        public enum ClassPlacement : int
        {
            /// <summary>
            /// A class field.
            /// </summary>
            FIELD,

            /// <summary>
            /// A class constructor.
            /// </summary>
            CONSTRUCTOR,

            /// <summary>
            /// A class finalizer.
            /// </summary>
            FINALIZER,

            /// <summary>
            /// A delegate.
            /// </summary>
            DELEGATE,

            /// <summary>
            /// An class' event.
            /// </summary>
            EVENT,

            /// <summary>
            /// An enumeration
            /// </summary>
            ENUM,

            /// <summary>
            /// An interface implemented within a class.
            /// </summary>
            INTERFACE,

            /// <summary>
            /// A class property.
            /// </summary>
            PROPERTY,

            /// <summary>
            /// A method within a class.
            /// </summary>
            METHOD,

            /// <summary>
            /// A struct within a class.
            /// </summary>
            STRUCT,

            /// <summary>
            /// A class within a class.
            /// </summary>
            CLASS
        }

        /// <summary>
        /// Gets or sets a value indicating whether the task is disabled for a single iteration.
        /// </summary>
        public bool IsTemporarilyDisabled { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the task is enabled or not.
        /// </summary>
        public override bool IsEnabled
        {
            get
            {
                return Properties.Settings.Default.SortElementsWithinClassEnabled;
            }

            set
            {
                Properties.Settings.Default.SortElementsWithinClassEnabled = value;
            }
        }

        /// <summary>
        /// Resets temporary disabling that would affect the remainder of a pass, but not the next pass.
        /// </summary>
        public override void Reset()
        {
            IsTemporarilyDisabled = false;
        }

        /// <summary>
        /// Performs the style task.
        /// </summary>
        /// <param name="projectItem">The project Item</param>
        /// <param name="ideWindow">The IDE window.</param>
        protected override void DoWork(ProjectItem projectItem, EnvDTE.Window ideWindow)
        {
            if (projectItem.Name.EndsWith(".cs") && !IsTemporarilyDisabled)
            {
                Debug.WriteLine("Sorting Functions: " + projectItem.Name);
                FileCodeModel fileCodeModel = projectItem.FileCodeModel;

                if (CheckForPreprocessorDirectives(ideWindow))
                {
                    if (MessageBox.Show("There appear to be preprocessor directives in this file.  Sorting elements within a class may break scope of these.  Do you want to continue the action on this file?", "Warning: Preprocessor Directives Detected", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) == DialogResult.No)
                    {
                        return;
                    }
                }

                for (int i = 1; i <= fileCodeModel.CodeElements.Count; i++)
                {
                    if (fileCodeModel.CodeElements.Item(i).Kind == vsCMElement.vsCMElementNamespace)
                    {
                        CodeElement element = fileCodeModel.CodeElements.Item(i);
                        for (int j = 1; j <= element.Children.Count; j++)
                        {
                            if (element.Children.Item(j).Kind == vsCMElement.vsCMElementClass)
                            {
                                if (!EvaluateElementsWithinClassSorted(element.Children.Item(j)))
                                {
                                    SortFunctionsWithinClass(element.Children.Item(j));
                                }
                            }
                            else
                            {
                                Debug.WriteLine("--" + element.Children.Item(j).Kind);
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Checks the document in the window for preprocessor directives.
        /// </summary>
        /// <param name="ideWindow">The window of the document to check.</param>
        /// <returns>True if there are preprocessor directives present.</returns>
        private bool CheckForPreprocessorDirectives(EnvDTE.Window ideWindow)
        {
            TextDocument doc = (TextDocument)ideWindow.Document.Object("TextDocument");
            EditPoint editPoint = doc.CreateEditPoint(doc.StartPoint);

            while (!editPoint.AtEndOfDocument)
            {
                if (editPoint.GetLines(editPoint.Line, editPoint.Line + 1).StartsWith("#"))
                {
                    return true;
                }

                editPoint.LineDown(1);
            }

            return false;
        }

        bool EvaluateElementsWithinClassSorted(CodeElement codeElement)
        {
            try
            {
                CodeBlock lastBlock = null;
                CodeBlock currentBlock = null;
                Array accessLevels = Enum.GetValues(typeof(vsCMAccess));
                for (int i = 1; i <= codeElement.Children.Count; i++)
                {
                    CodeElement element = codeElement.Children.Item(i);
                    EditPoint elementStartPoint = element.StartPoint.CreateEditPoint();
                    EditPoint newStartPoint = elementStartPoint.CreateEditPoint();
                    switch (element.Kind)
                    {
                        case vsCMElement.vsCMElementVariable:
                            CodeVariable variable = element as CodeVariable;
                            if (variable != null)
                            {
                                currentBlock = new CodeBlock(variable.Access, ClassPlacement.FIELD, GetCodeBlockText(codeElement, element, out newStartPoint));
                            }
                            else
                            {
                                Debug.WriteLine("CodeVariable " + element.Name + " null");
                            }

                            break;
                        case vsCMElement.vsCMElementFunction:
                            // method, constructor, or finalizer
                            CodeFunction function = element as CodeFunction;
                            if (function != null)
                            {
                                if (function.FunctionKind == vsCMFunction.vsCMFunctionConstructor)
                                {
                                    currentBlock = new CodeBlock(function.Access, ClassPlacement.CONSTRUCTOR, GetCodeBlockText(codeElement, element, out newStartPoint));
                                }
                                else if (function.FunctionKind == vsCMFunction.vsCMFunctionDestructor)
                                {
                                    currentBlock = new CodeBlock(function.Access, ClassPlacement.FINALIZER, GetCodeBlockText(codeElement, element, out newStartPoint));
                                }
                                else
                                {
                                    currentBlock = new CodeBlock(function.Access, ClassPlacement.METHOD, GetCodeBlockText(codeElement, element, out newStartPoint));
                                }
                            }
                            else
                            {
                                Debug.WriteLine("CodeFunction " + element.Name + " null");
                            }

                            break;
                        case vsCMElement.vsCMElementDelegate:
                            CodeDelegate delegateElement = element as CodeDelegate;
                            if (delegateElement != null)
                            {
                                currentBlock = new CodeBlock(delegateElement.Access, ClassPlacement.DELEGATE, GetCodeBlockText(codeElement, element, out newStartPoint));
                            }
                            else
                            {
                                Debug.WriteLine("CodeDelegate " + element.Name + " null");
                            }

                            break;
                        case vsCMElement.vsCMElementEvent:
                            currentBlock = new CodeBlock(vsCMAccess.vsCMAccessPublic, ClassPlacement.EVENT, GetCodeBlockText(codeElement, element, out newStartPoint));
                            break;
                        case vsCMElement.vsCMElementEnum:
                            CodeEnum enumElement = element as CodeEnum;
                            if (enumElement != null)
                            {
                                currentBlock = new CodeBlock(enumElement.Access, ClassPlacement.ENUM, GetCodeBlockText(codeElement, element, out newStartPoint));
                            }
                            else
                            {
                                Debug.WriteLine("CodeEnum " + element.Name + " null");
                            }

                            break;
                        case vsCMElement.vsCMElementInterface:
                            CodeInterface interfaceElement = element as CodeInterface;
                            if (interfaceElement != null)
                            {
                                currentBlock = new CodeBlock(interfaceElement.Access, ClassPlacement.INTERFACE, GetCodeBlockText(codeElement, element, out newStartPoint));
                            }
                            else
                            {
                                Debug.WriteLine("CodeInterface " + element.Name + " null");
                            }

                            break;
                        case vsCMElement.vsCMElementProperty:
                            CodeProperty propertyElement = element as CodeProperty;
                            if (propertyElement != null)
                            {
                                currentBlock = new CodeBlock(propertyElement.Access, ClassPlacement.PROPERTY, GetCodeBlockText(codeElement, element, out newStartPoint));
                            }
                            else
                            {
                                Debug.WriteLine("CodeProperty " + element.Name + " null");
                            }

                            break;
                        case vsCMElement.vsCMElementStruct:
                            CodeStruct structElement = element as CodeStruct;
                            if (structElement != null)
                            {
                                currentBlock = new CodeBlock(structElement.Access, ClassPlacement.STRUCT, GetCodeBlockText(codeElement, element, out newStartPoint));
                            }
                            else
                            {
                                Debug.WriteLine("CodeStruct " + element.Name + " null");
                            }

                            break;
                        case vsCMElement.vsCMElementClass:
                            CodeClass classElement = element as CodeClass;
                            if (classElement != null)
                            {
                                currentBlock = new CodeBlock(classElement.Access, ClassPlacement.CLASS, GetCodeBlockText(codeElement, element, out newStartPoint));
                            }
                            else
                            {
                                Debug.WriteLine("CodeStruct " + element.Name + " null");
                            }

                            break;
                        default:
                            Debug.WriteLine("unknown element: " + element.Name + " - " + element.Kind);
                            currentBlock = new CodeBlock(vsCMAccess.vsCMAccessPrivate, ClassPlacement.CLASS, "/*UKNOWN*/ " + GetCodeBlockText(codeElement, element, out newStartPoint));
                            break;
                    }

                    if (lastBlock != null)
                    {
                        if (lastBlock.Placement != currentBlock.Placement)
                        {
                            if (lastBlock.Placement.CompareTo(currentBlock.Placement) > 0)
                            {
                                Debug.WriteLine(currentBlock.Placement + " - " + currentBlock.Access + " belongs before " + lastBlock.Placement + " - " + lastBlock.Access + "; Sorting Required.");
                                return false;
                            }
                        }
                        else
                        {
                            if (lastBlock.Access.CompareTo(currentBlock.Access) > 0)
                            {
                                Debug.WriteLine(currentBlock.Placement + " - " + currentBlock.Access + " belongs before " + lastBlock.Placement + " - " + lastBlock.Access + "; Sorting Required.");
                                return false;
                            }
                        }
                    }
                    else
                    {
                        lastBlock = currentBlock;
                    }
                }
            }
            catch (Exception exc)
            {
                Debug.WriteLine(exc.ToString());
                Debug.WriteLine("-- Class Reverted --");
            }

            Debug.WriteLine("Class is already sorted, returning");
            return true;
        }


        /// <summary>
        /// Sorts functions within a class.
        /// </summary>
        /// <param name="codeElement">The code element that represents the class.</param>
        private void SortFunctionsWithinClass(CodeElement codeElement)
        {
            EditPoint classPoint = codeElement.StartPoint.CreateEditPoint();
            TextRanges trs = null;

            string classBackup = classPoint.GetText(codeElement.EndPoint);

            try
            {
                if (classPoint.FindPattern("{", (int)vsFindOptions.vsFindOptionsMatchCase, ref classPoint, ref trs))
                {
                    classPoint.Insert("\r\n");

                    List<CodeBlock> blocks = new List<CodeBlock>();
                    Array accessLevels = Enum.GetValues(typeof(vsCMAccess));
                    for (int i = 1; i <= codeElement.Children.Count; i++)
                    {
                        CodeElement element = codeElement.Children.Item(i);
                        EditPoint elementStartPoint = element.StartPoint.CreateEditPoint();
                        EditPoint newStartPoint = elementStartPoint.CreateEditPoint();
                        switch (element.Kind)
                        {
                            case vsCMElement.vsCMElementVariable:
                                CodeVariable variable = element as CodeVariable;
                                if (variable != null)
                                {
                                    blocks.Add(new CodeBlock(variable.Access, ClassPlacement.FIELD, GetCodeBlockText(codeElement, element, out newStartPoint)));
                                }
                                else
                                {
                                    Debug.WriteLine("CodeVariable " + element.Name + " null");
                                }

                                break;
                            case vsCMElement.vsCMElementFunction:
                                // method, constructor, or finalizer
                                CodeFunction function = element as CodeFunction;
                                if (function != null)
                                {
                                    if (function.FunctionKind == vsCMFunction.vsCMFunctionConstructor)
                                    {
                                        blocks.Add(new CodeBlock(function.Access, ClassPlacement.CONSTRUCTOR, GetCodeBlockText(codeElement, element, out newStartPoint)));
                                    }
                                    else if (function.FunctionKind == vsCMFunction.vsCMFunctionDestructor)
                                    {
                                        blocks.Add(new CodeBlock(function.Access, ClassPlacement.FINALIZER, GetCodeBlockText(codeElement, element, out newStartPoint)));
                                    }
                                    else
                                    {
                                        blocks.Add(new CodeBlock(function.Access, ClassPlacement.METHOD, GetCodeBlockText(codeElement, element, out newStartPoint)));
                                    }
                                }
                                else
                                {
                                    Debug.WriteLine("CodeFunction " + element.Name + " null");
                                }

                                break;
                            case vsCMElement.vsCMElementDelegate:
                                CodeDelegate delegateElement = element as CodeDelegate;
                                if (delegateElement != null)
                                {
                                    blocks.Add(new CodeBlock(delegateElement.Access, ClassPlacement.DELEGATE, GetCodeBlockText(codeElement, element, out newStartPoint)));
                                }
                                else
                                {
                                    Debug.WriteLine("CodeDelegate " + element.Name + " null");
                                }

                                break;
                            case vsCMElement.vsCMElementEvent:
                                blocks.Add(new CodeBlock(vsCMAccess.vsCMAccessPublic, ClassPlacement.EVENT, GetCodeBlockText(codeElement, element, out newStartPoint)));
                                break;
                            case vsCMElement.vsCMElementEnum:
                                CodeEnum enumElement = element as CodeEnum;
                                if (enumElement != null)
                                {
                                    blocks.Add(new CodeBlock(enumElement.Access, ClassPlacement.ENUM, GetCodeBlockText(codeElement, element, out newStartPoint)));
                                }
                                else
                                {
                                    Debug.WriteLine("CodeEnum " + element.Name + " null");
                                }

                                break;
                            case vsCMElement.vsCMElementInterface:
                                CodeInterface interfaceElement = element as CodeInterface;
                                if (interfaceElement != null)
                                {
                                    blocks.Add(new CodeBlock(interfaceElement.Access, ClassPlacement.INTERFACE, GetCodeBlockText(codeElement, element, out newStartPoint)));
                                }
                                else
                                {
                                    Debug.WriteLine("CodeInterface " + element.Name + " null");
                                }

                                break;
                            case vsCMElement.vsCMElementProperty:
                                CodeProperty propertyElement = element as CodeProperty;
                                if (propertyElement != null)
                                {
                                    blocks.Add(new CodeBlock(propertyElement.Access, ClassPlacement.PROPERTY, GetCodeBlockText(codeElement, element, out newStartPoint)));
                                }
                                else
                                {
                                    Debug.WriteLine("CodeProperty " + element.Name + " null");
                                }

                                break;
                            case vsCMElement.vsCMElementStruct:
                                CodeStruct structElement = element as CodeStruct;
                                if (structElement != null)
                                {
                                    blocks.Add(new CodeBlock(structElement.Access, ClassPlacement.STRUCT, GetCodeBlockText(codeElement, element, out newStartPoint)));
                                }
                                else
                                {
                                    Debug.WriteLine("CodeStruct " + element.Name + " null");
                                }

                                break;
                            case vsCMElement.vsCMElementClass:
                                CodeClass classElement = element as CodeClass;
                                if (classElement != null)
                                {
                                    blocks.Add(new CodeBlock(classElement.Access, ClassPlacement.CLASS, GetCodeBlockText(codeElement, element, out newStartPoint)));
                                }
                                else
                                {
                                    Debug.WriteLine("CodeStruct " + element.Name + " null");
                                }

                                break;
                            default:
                                Debug.WriteLine("unknown element: " + element.Name + " - " + element.Kind);
                                blocks.Add(new CodeBlock(vsCMAccess.vsCMAccessPrivate, ClassPlacement.CLASS, "/*UKNOWN*/ " + GetCodeBlockText(codeElement, element, out newStartPoint)));
                                break;
                        }

                        newStartPoint.Delete(element.EndPoint);
                        newStartPoint.DeleteWhitespace(vsWhitespaceOptions.vsWhitespaceOptionsVertical);

                        i--;
                    }

                    if (codeElement.Children.Count != 0)
                    {
                        System.Diagnostics.Debugger.Break();
                    }

                    blocks.Sort(delegate(CodeBlock c1, CodeBlock c2)
                    {
                        if (c1.Placement != c2.Placement)
                        {
                            return c1.Placement.CompareTo(c2.Placement);
                        }
                        else
                        {
                            return c1.Access.CompareTo(c2.Access);
                        }
                    });

                    classPoint.DeleteWhitespace(vsWhitespaceOptions.vsWhitespaceOptionsVertical);
                    classPoint.Insert("\r\n");

                    for (int i = 0; i < blocks.Count; i++)
                    {
                        classPoint.Insert(blocks[i].Body + "\r\n\r\n");
                    }

                    classPoint.LineUp(1);
                    classPoint.DeleteWhitespace(vsWhitespaceOptions.vsWhitespaceOptionsVertical);

                    for (int i = 1; i <= codeElement.Children.Count; i++)
                    {
                        CodeElement element = codeElement.Children.Item(i);
                        if (element.Kind == vsCMElement.vsCMElementClass || element.Kind == vsCMElement.vsCMElementInterface || element.Kind == vsCMElement.vsCMElementStruct)
                        {
                            SortFunctionsWithinClass(element);
                        }
                    }
                }
            }
            catch (Exception exc)
            {
                Debug.WriteLine(exc.ToString());
                EditPoint startBackup = codeElement.StartPoint.CreateEditPoint();
                startBackup.Delete(codeElement.EndPoint);
                startBackup.Insert(classBackup);
                Debug.WriteLine("-- Class Reverted --");
            }
        }

        /// <summary>
        /// Gets the code block for a given element, this includes any comments that come before the element, 
        /// but after the previous element.
        /// </summary>
        /// <param name="parent">The current block's parent.</param>
        /// <param name="block">The current block.</param>
        /// <param name="newStartPoint">The starting point used to capture the beginning of the element.</param>
        /// <returns>The text of a given block.</returns>
        private string GetCodeBlockText(CodeElement parent, CodeElement block, out EditPoint newStartPoint)
        {
            EditPoint blockStartPoint = block.StartPoint.CreateEditPoint();
            EditPoint previousblockEndPoint = blockStartPoint.CreateEditPoint();
            EditPoint editPoint = blockStartPoint.CreateEditPoint();
            TextRanges trs = null;

            if (editPoint.FindPattern("}", (int)vsFindOptions.vsFindOptionsBackwards, ref previousblockEndPoint, ref trs))
            {
                if (previousblockEndPoint.GreaterThan(parent.StartPoint))
                {
                    blockStartPoint = previousblockEndPoint;
                }
                else
                {
                    previousblockEndPoint = blockStartPoint.CreateEditPoint();
                    editPoint = blockStartPoint.CreateEditPoint();
                    if (editPoint.FindPattern("{", (int)vsFindOptions.vsFindOptionsBackwards, ref previousblockEndPoint, ref trs))
                    {
                        if (previousblockEndPoint.GreaterThan(parent.StartPoint))
                        {
                            blockStartPoint = previousblockEndPoint;
                        }
                    }
                }
            }
            else
            {
                if (editPoint.FindPattern("{", (int)vsFindOptions.vsFindOptionsBackwards, ref previousblockEndPoint, ref trs))
                {
                    if (previousblockEndPoint.GreaterThan(parent.StartPoint))
                    {
                        blockStartPoint = previousblockEndPoint;
                    }
                }
            }

            newStartPoint = blockStartPoint.CreateEditPoint();

            return blockStartPoint.GetText(block.EndPoint).Trim();
        }

        /// <summary>
        /// Class representing a code block for storing and sorting.
        /// </summary>
        private class CodeBlock
        {
            /// <summary>
            /// Initializes a new instance of the CodeBlock class.
            /// </summary>
            /// <param name="access">The block's access level.</param>
            /// <param name="placement">The block's placement.</param>
            /// <param name="body">The body of the code block.</param>
            public CodeBlock(vsCMAccess access, ClassPlacement placement, string body)
            {
                Access = access;
                Body = body;
                Placement = placement;
            }

            /// <summary>
            /// Gets the block's access level.
            /// </summary>
            public vsCMAccess Access { get; private set; }

            /// <summary>
            /// Gets the block's body text.
            /// </summary>
            public string Body { get; private set; }

            /// <summary>
            /// Gets the block's placement level 
            /// </summary>
            public ClassPlacement Placement { get; private set; }
        }
    }
}