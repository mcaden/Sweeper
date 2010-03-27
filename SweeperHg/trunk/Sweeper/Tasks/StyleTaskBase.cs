namespace Sweeper
{
    using EnvDTE;

    /// <summary>
    /// Base class for Style tasks
    /// </summary>
    public abstract class StyleTaskBase : IProjectItemStyleTask
    {
        /// <summary>
        /// Gets or sets the name of the style task.
        /// </summary>
        public string TaskName { get; protected set; }

        /// <summary>
        /// Gets or sets a description of the style task.
        /// </summary>
        public string TaskDescription { get; protected set; }

        /// <summary>
        /// Gets or sets a value indicating whether the task is enabled or not.
        /// </summary>
        public abstract bool IsEnabled { get; set; }

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
