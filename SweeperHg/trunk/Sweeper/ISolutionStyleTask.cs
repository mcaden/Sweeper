namespace Sweeper
{
    /// <summary>
    /// Interface for varying tasks to be performed.
    /// </summary>
    public interface ISolutionStyleTask : IStyleTask
    {
        /// <summary>
        /// Performs the style task.
        /// </summary>
        /// <param name="ideWindow">The IDE window.</param>
        void PerformStyleTask(EnvDTE.Window ideWindow);
    }
}
