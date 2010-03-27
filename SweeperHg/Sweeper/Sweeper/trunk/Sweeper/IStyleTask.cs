namespace Sweeper
{
    /// <summary>
    /// Interface for varying tasks to be performed.
    /// </summary>
    public interface IStyleTask
    {
        /// <summary>
        /// Gets the name of the style task.
        /// </summary>
        string TaskName { get; }

        /// <summary>
        /// Gets a description of the style task.
        /// </summary>
        string TaskDescription { get; }

        /// <summary>
        /// Gets or sets a value indicating whether the task is enabled or not.
        /// </summary>
        bool IsEnabled { get; set; }

        /// <summary>
        /// Resets temporary disabling that would affect a complete pass.
        /// </summary>
        void Reset();
    }
}
