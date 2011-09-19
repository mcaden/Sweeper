namespace Sweeper
{
    using System;
    using System.Collections.Generic;
    using EnvDTE;
    using EnvDTE80;

    /// <summary>
    /// Base class for Style tasks
    /// </summary>
    public abstract class StyleTaskBase : IProjectItemStyleTask
    {
        /// <summary>
        /// Keywords used for code element access
        /// </summary>
        public static readonly Dictionary<vsCMAccess, string> CodeAccessKeywords = new Dictionary<vsCMAccess, string>() 
        { 
                { vsCMAccess.vsCMAccessPublic, "public" }, 
                { vsCMAccess.vsCMAccessProtected, "protected" },
                { vsCMAccess.vsCMAccessPrivate, "private" },
                { vsCMAccess.vsCMAccessProject, "internal" }
        };

        /// <summary>
        /// Element types used for code element access
        /// </summary>
        public static readonly Dictionary<vsCMElement, ElementType> CodeElementBlockTypes = new Dictionary<vsCMElement, ElementType>() 
        { 
                { vsCMElement.vsCMElementClass, ElementType.CLASS }, 
                { vsCMElement.vsCMElementDelegate, ElementType.DELEGATE },
                { vsCMElement.vsCMElementEnum, ElementType.ENUM },
                { vsCMElement.vsCMElementEvent, ElementType.EVENT },
                { vsCMElement.vsCMElementFunction, ElementType.METHOD },
                { vsCMElement.vsCMElementProperty, ElementType.PROPERTY },
                { vsCMElement.vsCMElementStruct, ElementType.STRUCT },
                { vsCMElement.vsCMElementVariable, ElementType.FIELD },
                { vsCMElement.vsCMElementInterface, ElementType.INTERFACE }
        };

        /// <summary>
        /// Keywords used for code element access
        /// </summary>
        public static readonly Dictionary<vsCMElement, Type> CodeElementTypes = new Dictionary<vsCMElement, Type>() 
        { 
                { vsCMElement.vsCMElementClass, typeof(CodeClass) }, 
                { vsCMElement.vsCMElementDelegate, typeof(CodeDelegate) },
                { vsCMElement.vsCMElementEnum, typeof(CodeEnum) },
                { vsCMElement.vsCMElementEvent, typeof(CodeEvent) },
                { vsCMElement.vsCMElementFunction, typeof(CodeFunction) },
                { vsCMElement.vsCMElementProperty, typeof(CodeProperty) },
                { vsCMElement.vsCMElementStruct, typeof(CodeStruct) },
                { vsCMElement.vsCMElementVariable, typeof(CodeVariable) },
                { vsCMElement.vsCMElementInterface, typeof(CodeInterface) }
        };

        /// <summary>
        /// Class priority enums to sort elements by their type.  
        /// Built-in type enums are not sufficient for this because they span multiple types.
        /// </summary>
        public enum ElementType : int
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
            /// A class' event.
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
        /// Gets or sets a value indicating whether the task is enabled or not.
        /// </summary>
        public abstract bool IsEnabled { get; set; }

        /// <summary>
        /// Gets or sets a description of the style task.
        /// </summary>
        public string TaskDescription { get; protected set; }

        /// <summary>
        /// Gets or sets the name of the style task.
        /// </summary>
        public string TaskName { get; protected set; }

        /// <summary>
        /// Performs the style task.
        /// </summary>
        /// <param name="projectItem">The project Item</param>
        /// <param name="ideWindow">The IDE window.</param>
        public void PerformStyleTask(ProjectItem projectItem, EnvDTE.Window ideWindow)
        {
            if (IsEnabled)
            {
                DoWork(projectItem, ideWindow);
            }
        }

        /// <summary>
        /// Resets temporary disabling that would affect the remainder of a pass, but not the next pass.
        /// </summary>
        public virtual void Reset()
        {
        }

        /// <summary>
        /// Performs the work within style task.
        /// </summary>
        /// <param name="projectItem">The project Item</param>
        /// <param name="ideWindow">The IDE window.</param>
        protected abstract void DoWork(ProjectItem projectItem, EnvDTE.Window ideWindow);
    }
}
