namespace Sweeper
{
    using System.Diagnostics;
    using EnvDTE;
    using System.Collections.Generic;
    using System;

    public class SortFunctions : StyleTaskBase
    {
        public SortFunctions()
        {
            TaskName = "Sort Functions";
            TaskDescription = "Sorts functions in the order of public, internal, protected, and private.";
            Debug.WriteLine("Task: " + TaskName + " created.");
        }

        /// <summary>
        /// Gets or sets a value indicating whether the task is enabled or not.
        /// </summary>
        public override bool IsEnabled
        {
            get
            {
                return Properties.Settings.Default.SortFunctions;
            }

            set
            {
                Properties.Settings.Default.SortFunctions = value;
            }
        }

        /// <summary>
        /// Performs the style task.
        /// </summary>
        /// <param name="projectItem">The project Item</param>
        protected override void DoWork(ProjectItem projectItem, EnvDTE.Window ideWindow)
        {
            if (projectItem.Name.EndsWith(".cs") && !projectItem.Name.EndsWith("Designer.cs"))
            {
                Debug.WriteLine("Sorting Functions: " + projectItem.Name);
                FileCodeModel fileCodeModel = projectItem.FileCodeModel;
                for (int i = 1; i <= fileCodeModel.CodeElements.Count; i++)
                {
                    if (fileCodeModel.CodeElements.Item(i).Kind == vsCMElement.vsCMElementNamespace)
                    {
                        CodeElement element = fileCodeModel.CodeElements.Item(i);
                        for (int j = 1; j <= element.Children.Count; j++)
                        {
                            if (element.Children.Item(j).Kind == vsCMElement.vsCMElementClass)
                            {
                                SortFunctionsWithinClass(element.Children.Item(j));
                            }
                        }
                    }
                }
            }
        }

        private void SortFunctionsWithinClass(CodeElement codeElement)
        {
            EditPoint fieldInsertionPoint;
            EditPoint constructorInsertionPoint;
            EditPoint finalizerInsertionPoint;
            EditPoint delegateInsertionPoint;
            EditPoint eventInsertionPoint;
            EditPoint enumInsertionPoint;
            EditPoint interfaceInsertionPoint;
            EditPoint propertyInsertionPoint;
            EditPoint methodInsertionPoint;
            EditPoint structInsertionPoint;
            EditPoint classInsertionPoint;

            EditPoint origin = codeElement.StartPoint.CreateEditPoint();
            EditPoint classPoint = codeElement.StartPoint.CreateEditPoint();
            TextRanges trs = null;

            if (origin.FindPattern("{", (int)vsFindOptions.vsFindOptionsMatchCase, ref classPoint, ref trs))
            {
                classPoint.Insert("\r\n");
                origin = classPoint.CreateEditPoint();
                fieldInsertionPoint = classPoint.CreateEditPoint();
                classPoint.Insert("\r\n\r\n");
                constructorInsertionPoint = classPoint.CreateEditPoint();
                classPoint.Insert("\r\n\r\n");
                finalizerInsertionPoint = classPoint.CreateEditPoint();
                classPoint.Insert("\r\n\r\n");
                delegateInsertionPoint = classPoint.CreateEditPoint();
                classPoint.Insert("\r\n\r\n");
                eventInsertionPoint = classPoint.CreateEditPoint();
                classPoint.Insert("\r\n\r\n");
                enumInsertionPoint = classPoint.CreateEditPoint();
                classPoint.Insert("\r\n\r\n");
                interfaceInsertionPoint = classPoint.CreateEditPoint();
                classPoint.Insert("\r\n\r\n");
                propertyInsertionPoint = classPoint.CreateEditPoint();
                classPoint.Insert("\r\n\r\n");
                methodInsertionPoint = classPoint.CreateEditPoint();
                classPoint.Insert("\r\n\r\n");
                structInsertionPoint = classPoint.CreateEditPoint();
                classPoint.Insert("\r\n\r\n");
                classInsertionPoint = classPoint.CreateEditPoint();
                classPoint.Insert("\r\n\r\n");

                Array accessLevels = Enum.GetValues(typeof(vsCMAccess));

                foreach (vsCMAccess accessLevel in accessLevels)
                {
                    for (int i = 1; i <= codeElement.Children.Count; i++)
                    {
                        CodeElement element = codeElement.Children.Item(i);
                        EditPoint elementStartPoint = element.StartPoint.CreateEditPoint();
                        switch (element.Kind)
                        {
                            case vsCMElement.vsCMElementVariable:
                                CodeVariable variable = element as CodeVariable;
                                if (variable != null)
                                {
                                    if (variable.Access == accessLevel)
                                    {
                                        MoveCodeBlock(codeElement, element, fieldInsertionPoint);
                                    }
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
                                    if (function.Access == accessLevel)
                                    {
                                        if (function.FunctionKind == vsCMFunction.vsCMFunctionConstructor)
                                        {
                                            MoveCodeBlock(codeElement, element, constructorInsertionPoint);
                                        }
                                        else if (function.FunctionKind == vsCMFunction.vsCMFunctionDestructor)
                                        {
                                            MoveCodeBlock(codeElement, element, finalizerInsertionPoint);
                                        }
                                        else
                                        {
                                            MoveCodeBlock(codeElement, element, methodInsertionPoint);
                                        }
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
                                    if (delegateElement.Access == accessLevel)
                                    {
                                        MoveCodeBlock(codeElement, element, delegateInsertionPoint);
                                    }
                                }
                                else
                                {
                                    Debug.WriteLine("CodeDelegate " + element.Name + " null");
                                }
                                break;
                            case vsCMElement.vsCMElementEvent:
                                MoveCodeBlock(codeElement, element, eventInsertionPoint);
                                break;
                            case vsCMElement.vsCMElementEnum:
                                CodeEnum enumElement = element as CodeEnum;
                                if (enumElement != null)
                                {
                                    if (enumElement.Access == accessLevel)
                                    {
                                        MoveCodeBlock(codeElement, element, enumInsertionPoint);
                                    }
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
                                    if (interfaceElement.Access == accessLevel)
                                    {
                                        MoveCodeBlock(codeElement, element, interfaceInsertionPoint);
                                    }
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
                                    if (propertyElement.Access == accessLevel)
                                    {
                                        MoveCodeBlock(codeElement, element, propertyInsertionPoint);
                                    }
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
                                    if (structElement.Access == accessLevel)
                                    {
                                        MoveCodeBlock(codeElement, element, structInsertionPoint);
                                    }
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
                                    if (classElement.Access == accessLevel)
                                    {
                                        MoveCodeBlock(codeElement, element, classInsertionPoint);
                                    }
                                }
                                else
                                {
                                    Debug.WriteLine("CodeStruct " + element.Name + " null");
                                }
                                break;
                            default:
                                Debug.WriteLine("unknown element: " + element.Name + " - " + element.Kind);
                                break;
                        }
                    }

                    for (int i = 1; i <= codeElement.Children.Count; i++)
                    {
                        CodeElement element = codeElement.Children.Item(i);
                        if (element.Kind == vsCMElement.vsCMElementClass || element.Kind == vsCMElement.vsCMElementInterface || element.Kind == vsCMElement.vsCMElementStruct)
                        {
                            SortFunctionsWithinClass(element);
                        }
                    }
                }

                origin.DeleteWhitespace(vsWhitespaceOptions.vsWhitespaceOptionsVertical);
                classInsertionPoint.CreateEditPoint().DeleteWhitespace(vsWhitespaceOptions.vsWhitespaceOptionsVertical);
            }
        }

        private static void MoveCodeBlock(CodeElement parent, CodeElement block, EditPoint pastePoint)
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
            }

            blockStartPoint.Cut(block.EndPoint, false);
            pastePoint.Paste();
            pastePoint.Insert("\r\n\r\n");
        }
    }
}