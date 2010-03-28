namespace SweeperUI
{
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.Windows;

    /// <summary>
    /// Interaction logic for OptionWindow.xaml
    /// </summary>
    public partial class OptionWindow : Window
    {
        /// <summary>
        /// Initializes a new instance of the OptionWindow class.
        /// </summary>
        /// <param name="newOptions">The options to display.</param>
        public OptionWindow(List<object> newOptions)
        {
            Options = new ObservableCollection<object>(newOptions);
            this.DataContext = this;
            InitializeComponent();
        }

        /// <summary>
        /// Gets or sets the options displayed.
        /// </summary>
        public ObservableCollection<object> Options { get; set; }

        /// <summary>
        /// Handles the button click event for "OK".
        /// </summary>
        /// <param name="sender">The sender object.</param>
        /// <param name="e">Event arguments.</param>
        private void OK_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = true;
            this.Close();
        }
    }
}
