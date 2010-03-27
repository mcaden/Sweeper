namespace Sweeper
{
    using EnvDTE;

    /// <summary>
    /// Interface for varying tasks to be performed.
    /// </summary>
    public interface IProjectItemStyleTask : IStyleTask
    {
        /// <summary>
        /// Performs the style task.
        /// </summary>
        /// <param name="projectItem">The project Item</param>
        /// <param name="ideWindow">The IDE window.</param>
        void PerformStyleTask(ProjectItem projectItem, EnvDTE.Window ideWindow);
    }
}
