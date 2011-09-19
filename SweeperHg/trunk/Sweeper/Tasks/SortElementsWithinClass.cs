namespace Sweeper
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.Reflection;
    using System.Windows.Forms;
    using EnvDTE;
    using EnvDTE80;

    /// <summary>
    /// Sorts Elements according to their type and access level to fit standard.
    /// </summary>
    public class SortElementsWithinClass : StyleTaskBase
    {
        /// <summary>
        /// Lookup table between ElementAccess and visual studio's access enum
        /// </summary>
        private readonly Dictionary<vsCMAccess, ElementAccess> accessLookup = new Dictionary<vsCMAccess, ElementAccess>();

        /// <summary>
        /// Initializes a new instance of the SortElementsWithinClass class.
        /// </summary>
        public SortElementsWithinClass()
        {
            TaskName = "Sort Elements Within a Class";
            TaskDescription = "Sorts Elements according to their type and access level to fit standard.";
            Debug.WriteLine("Task: " + TaskName + " created.");
            IsTemporarilyDisabled = false;

            accessLookup.Add(vsCMAccess.vsCMAccessPrivate, ElementAccess.PRIVATE);
            accessLookup.Add(vsCMAccess.vsCMAccessProtected, ElementAccess.PROTECTED);
            accessLookup.Add(vsCMAccess.vsCMAccessProject, ElementAccess.INTERNAL);
            accessLookup.Add(vsCMAccess.vsCMAccessPublic, ElementAccess.PUBLIC);
        }

        /// <summary>
        /// An element's access type used to define order
        /// </summary>
        public enum ElementAccess : int
        {
            /// <summary>
            /// Public access
            /// </summary>
            PUBLIC_CONSTANT_STATIC,

            /// <summary>
            /// Public access
            /// </summary>
            PUBLIC_CONSTANT,

            /// <summary>
            /// Public access
            /// </summary>
            PUBLIC_STATIC,

            /// <summary>
            /// Public access
            /// </summary>
            PUBLIC,

            /// <summary>
            /// Internal to project
            /// </summary>
            INTERNAL_CONSTANT_STATIC,

            /// <summary>
            /// Internal to project
            /// </summary>
            INTERNAL_CONSTANT,

            /// <summary>
            /// Internal to project
            /// </summary>
            INTERNAL_STATIC,

            /// <summary>
            /// Internal to project
            /// </summary>
            INTERNAL,

            /// <summary>
            /// Protected access
            /// </summary>
            PROTECTED_CONSTANT_STATIC,

            /// <summary>
            /// Protected access
            /// </summary>
            PROTECTED_CONSTANT,

            /// <summary>
            /// Protected access
            /// </summary>
            PROTECTED_STATIC,

            /// <summary>
            /// Protected access
            /// </summary>
            PROTECTED,

            /// <summary>
            /// Private access
            /// </summary>
            PRIVATE_CONSTANT_STATIC,

            /// <summary>
            /// Private access
            /// </summary>
            PRIVATE_CONSTANT,

            /// <summary>
            /// Private access
            /// </summary>
            PRIVATE_STATIC,

            /// <summary>
            /// Private access
            /// </summary>
            PRIVATE
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
                            vsCMElement kind = element.Children.Item(j).Kind;
                            if (kind == vsCMElement.vsCMElementClass || kind == vsCMElement.vsCMElementInterface || kind == vsCMElement.vsCMElementStruct)
                            {
                                if (!EvaluateElementsWithinClassSorted(element.Children.Item(j)))
                                {
                                    SortFunctionsWithinClass(element.Children.Item(j));
                                }
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

        /// <summary>
        /// Evaluates whether elements within a class have been sorted already or not.
        /// </summary>
        /// <param name="codeElement">The code element representing the class.</param>
        /// <returns>True if the class has already been sorted.</returns>
        private bool EvaluateElementsWithinClassSorted(CodeElement codeElement)
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

                    currentBlock = null;
                    if (CodeElementBlockTypes.ContainsKey(element.Kind))
                    {
                        currentBlock = EvaluateBlock(codeElement, element, ref newStartPoint);
                    }

                    if (currentBlock != null)
                    {
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
        /// Evaluates an element into a codeblock
        /// </summary>
        /// <param name="parentElement">The parent element, containing the given element</param>
        /// <param name="element">The element to evaluate</param>
        /// <param name="newStartPoint">The starting point for the element</param>
        /// <returns>A CodeBlock derived from the Element with appropriate placement and access set</returns>
        private CodeBlock EvaluateBlock(CodeElement parentElement, CodeElement element, ref EditPoint newStartPoint)
        {
            CodeBlock currentBlock = null;
            ElementType blockType = GetType(element);
            bool blockStatic = GetElementIsStatic(element);
            bool blockConstant = GetElementIsConstant(element);
            string body = GetCodeBlockText(parentElement, element, out newStartPoint);
            ElementAccess blockAccess = GetAccess(element, blockConstant, blockStatic);
            int weight = GetWeight(element);

            currentBlock = new CodeBlock(blockAccess, blockType, element.Name, body);
            return currentBlock;
        }

        /// <summary>
        /// Gets a weight used for sorting from a given element
        /// </summary>
        /// <param name="element">The current code element</param>
        /// <returns>A weight used for sorting from a given element</returns>
        private int GetWeight(CodeElement element)
        {
            int weight = 0;
            if (element.Kind == vsCMElement.vsCMElementFunction)
            {
                CodeFunction functionElement = (CodeFunction)element;

                foreach (CodeParameter2 param in functionElement.Parameters)
                {
                    string paramKind = param.ParameterKind.ToString().Replace("vsCMParameterKind", string.Empty);
                    paramKind = paramKind == "None" ? string.Empty : paramKind;
                    paramKind = paramKind == "ParamArray" ? "params" : paramKind;
                    weight += param.Type.AsString.Length + param.FullName.Length + paramKind.Length;
                }
            }

            return weight;
        }

        /// <summary>
        /// Gets a value indicating whether a given code element is static
        /// </summary>
        /// <param name="element">The current code element</param>
        /// <returns>A value indicating whether a given code element is static</returns>
        private bool GetElementIsStatic(CodeElement element)
        {
            object isStatic;
            bool isRetrieved = false;
            if (element.Kind == vsCMElement.vsCMElementProperty)
            {
                isRetrieved = GetElementProperty(element, "Getter.IsShared", out isStatic) || GetElementProperty(element, "Setter.IsShared", out isStatic);
            }
            else
            {
                isRetrieved = GetElementProperty(element, "IsShared", out isStatic);
            }

            if (!isRetrieved)
            {
                isStatic = false;
            }

            return (bool)isStatic;
        }

        /// <summary>
        /// Gets a value indicating whether a given code element is a constant
        /// </summary>
        /// <param name="element">The current code element</param>
        /// <returns>A value indicating whether a given code element is a constant</returns>
        private bool GetElementIsConstant(CodeElement element)
        {
            object isConstant;
            return GetElementProperty(element, "IsConstant", out isConstant) && (bool)isConstant;
        }

        /// <summary>
        /// Gets a property within a given element. Will search down the chain if given a compound path
        /// </summary>
        /// <param name="element">The current code element</param>
        /// <param name="propertyName">The property name to look for</param>
        /// <param name="propertyValue">The property value retrieved</param>
        /// <returns>True if the given property was retrieved</returns>
        private bool GetElementProperty(CodeElement element, string propertyName, out object propertyValue)
        {
            bool isSuccessful = false;
            propertyValue = null;
            if (!string.IsNullOrEmpty(propertyName) && propertyName.Trim() != string.Empty)
            {
                string[] propertyParts = propertyName.Split('.');
                PropertyInfo returnedProperty = CodeElementTypes[element.Kind].GetProperty(propertyParts[0]);
                if (returnedProperty != null)
                {
                    propertyValue = CodeElementTypes[element.Kind].InvokeMember(returnedProperty.Name, BindingFlags.GetProperty, null, element, null);
                }

                if (returnedProperty != null && (propertyValue != null && propertyParts.Length > 1))
                {
                    int index = 1;
                    while (propertyValue != null && returnedProperty != null && index < propertyParts.Length)
                    {
                        returnedProperty = returnedProperty.PropertyType.GetProperty(propertyParts[index]);
                        if (returnedProperty != null)
                        {
                            propertyValue = returnedProperty.ReflectedType.InvokeMember(returnedProperty.Name, BindingFlags.GetProperty, null, propertyValue, null);
                        }

                        index++;
                    }

                    if (index == propertyParts.Length && returnedProperty != null)
                    {
                        isSuccessful = true;
                    }
                }
                else if (returnedProperty != null && propertyParts.Length == 1)
                {
                    isSuccessful = true;
                }
            }

            if (!isSuccessful)
            {
                propertyValue = null;
            }

            return isSuccessful;
        }

        /// <summary>
        /// Gets the ElementType of a given code element
        /// </summary>
        /// <param name="element">The current code element</param>
        /// <returns>The ElementType of a given code element</returns>
        private ElementType GetType(CodeElement element)
        {
            ElementType type = CodeElementBlockTypes[element.Kind];
            if (element.Kind == vsCMElement.vsCMElementFunction)
            {
                CodeFunction2 functionElement = (CodeFunction2)element;
                if (functionElement.FunctionKind == vsCMFunction.vsCMFunctionConstructor)
                {
                    type = ElementType.CONSTRUCTOR;
                }
                else if (functionElement.FunctionKind == vsCMFunction.vsCMFunctionDestructor)
                {
                    type = ElementType.FINALIZER;
                }
            }

            return type;
        }

        /// <summary>
        /// Retrieves the access modifiers of an element
        /// </summary>
        /// <param name="element">The current code element</param>
        /// <param name="isConstant">Whether the element is a constant</param>
        /// <param name="isStatic">Whether the element is static</param>
        /// <returns>A valid ElementAccess</returns>
        private ElementAccess GetAccess(CodeElement element, bool isConstant, bool isStatic)
        {
            object access;
            if (!GetElementProperty(element, "Access", out access))
            {
                access = vsCMAccess.vsCMAccessPrivate;
            }

            return EvaluateAccess((vsCMAccess)access, isConstant, isStatic);
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
                        CodeBlock block = EvaluateBlock(codeElement, element, ref newStartPoint);

                        if (block != null)
                        {
                            blocks.Add(block);
                            newStartPoint.Delete(element.EndPoint);
                            newStartPoint.DeleteWhitespace(vsWhitespaceOptions.vsWhitespaceOptionsVertical);

                            i--;
                        }
                    }

                    blocks.Sort(delegate(CodeBlock c1, CodeBlock c2)
                    {
                        int comparison = 0;
                        if (c1.Placement != c2.Placement)
                        {
                            comparison = c1.Placement.CompareTo(c2.Placement);
                        }
                        else if (c1.Access != c2.Access)
                        {
                            comparison = c1.Access.CompareTo(c2.Access);
                        }
                        else if (c1.Name != c2.Name)
                        {
                            comparison = c1.Name.CompareTo(c2.Name);
                        }
                        else
                        {
                            comparison = c1.Weight.CompareTo(c2.Weight);
                        }

                        return comparison;
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
        /// Evaluates the appropriate access to use
        /// </summary>
        /// <param name="access">The visual studio access enum</param>
        /// <param name="isConstant">Whether the element is constant</param>
        /// <param name="isStatic">Whether the element is static</param>
        /// <returns>The appropriate access used for ordering</returns>
        private ElementAccess EvaluateAccess(vsCMAccess access, bool isConstant, bool isStatic)
        {
            ElementAccess elementAccess = accessLookup[access];
            if (isConstant)
            {
                elementAccess = (ElementAccess)Enum.Parse(typeof(ElementAccess), elementAccess.ToString() + "_CONSTANT");
            }

            if (isStatic)
            {
                elementAccess = (ElementAccess)Enum.Parse(typeof(ElementAccess), elementAccess.ToString() + "_STATIC");
            }

            return elementAccess;
        }

        /// <summary>
        /// Evaluates the appropriate access to use
        /// </summary>
        /// <param name="access">The visual studio access enum</param>
        /// <param name="isStatic">Whether the element is static</param>
        /// <returns>The appropriate access used for ordering</returns>
        private ElementAccess EvaluateAccess(vsCMAccess access, bool isStatic)
        {
            return EvaluateAccess(access, false, isStatic);
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
            /// <param name="name">The name of the block</param>
            /// <param name="body">The body of the code block.</param>
            public CodeBlock(ElementAccess access, ElementType placement, string name, string body)
            {
                Name = name;
                Access = access;
                Body = body;
                Placement = placement;
                Weight = 0;
            }

            /// <summary>
            /// Initializes a new instance of the CodeBlock class.
            /// </summary>
            /// <param name="access">The block's access level.</param>
            /// <param name="placement">The block's placement.</param>
            /// <param name="name">The name of the block</param>
            /// <param name="body">The body of the code block.</param>
            /// <param name="weight">An additional sort component that only applies to items that are the same access and placement.</param>
            public CodeBlock(ElementAccess access, ElementType placement, string name, string body, int weight)
                : this(access, placement, name, body)
            {
                Weight = weight;
            }

            /// <summary>
            /// Gets the block's access level.
            /// </summary>
            public ElementAccess Access { get; private set; }

            /// <summary>
            /// Gets the block's body text.
            /// </summary>
            public string Body { get; private set; }

            /// <summary>
            /// Gets the block's name.
            /// </summary>
            public string Name { get; private set; }

            /// <summary>
            /// Gets an additional sort component that only applies to items that are the same access and placement.
            /// </summary>
            public int Weight { get; private set; }

            /// <summary>
            /// Gets the block's placement level 
            /// </summary>
            public ElementType Placement { get; private set; }
        }
    }
}