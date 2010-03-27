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
        public OptionWindow(List<object> newOptions)
        {
            Options = new ObservableCollection<object>(newOptions);
            this.DataContext = this;
            InitializeComponent();
        }

        public ObservableCollection<object> Options { get; set; }

        private void OK_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = true;
            this.Close();
        }
    }
}
