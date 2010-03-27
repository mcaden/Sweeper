namespace Sweeper
{
    /// <summary>
    /// Base class for Style tasks
    /// </summary>
    public abstract class SolutionStyleTaskBase : ISolutionStyleTask
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
        /// <param name="ideWindow">The IDE window.</param>
        public void PerformStyleTask(EnvDTE.Window ideWindow)
        {
            if (IsEnabled)
            {
                DoWork(ideWindow);
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
        /// <param name="ideWindow">The IDE window.</param>
        protected abstract void DoWork(EnvDTE.Window ideWindow);
    }
}
