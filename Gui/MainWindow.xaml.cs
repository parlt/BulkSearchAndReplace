using System.Windows;
using System.Windows.Forms;
using BulkSearchAndReplaceLib;

namespace Gui
{
    /// <summary>
    ///     Interaktionslogik für MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        /// <summary>
        /// </summary>
        /// <param name="sender">object</param>
        /// <param name="e">RoutedEventArgs</param>
        private void SelectConfigurationFile_OnClick(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = BulkSearchAndReplace.FileFilterTxt;
            if (openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                SelectConfigurationFileResult.Content = openFileDialog.SafeFileName;
                BulkSearchAndReplace.GetInstance().ConfigFilePath = openFileDialog.FileName;
            }
        }

        /// <summary>
        /// </summary>
        /// <param name="sender">object</param>
        /// <param name="e">RoutedEventArgs</param>
        private void SelectExcelFile_OnClick(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = BulkSearchAndReplace.FileFilterExcel;
            if (openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                SelectExcelFileResult.Content = openFileDialog.SafeFileName;
                BulkSearchAndReplace.GetInstance().ExcelFilePath = openFileDialog.FileName;
            }
        }

        /// <summary>
        /// </summary>
        /// <param name="sender">object</param>
        /// <param name="e">RoutedEventArgs</param>
        private void SelectSourceDirectory_OnClick(object sender, RoutedEventArgs e)
        {
            var dialog = new FolderBrowserDialog();
            var result = dialog.ShowDialog();

            if (result == System.Windows.Forms.DialogResult.OK)
            {
                SelectSourceDirectoryResult.Content = dialog.SelectedPath;
                BulkSearchAndReplace.GetInstance().SourceDirectorPath = dialog.SelectedPath;
            }
        }

        /// <summary>
        /// </summary>
        /// <param name="sender">object</param>
        /// <param name="e">RoutedEventArgs</param>
        private void SelectDestinationDirectory_OnClick(object sender, RoutedEventArgs e)
        {
            var dialog = new FolderBrowserDialog();
            var result = dialog.ShowDialog();
            if (result == System.Windows.Forms.DialogResult.OK)
            {
                SelectDestinationDirectoryResult.Content = dialog.SelectedPath;
                BulkSearchAndReplace.GetInstance().DestinationDirectoryPath = dialog.SelectedPath;
            }
        }

        /// <summary>
        /// </summary>
        /// <param name="sender">object</param>
        /// <param name="e">RoutedEventArgs</param>
        private void SelectRun_OnClick(object sender, RoutedEventArgs e)
        {
            var message = BulkSearchAndReplace.GetInstance().Run();
            Message.Content = message;
        }
    }
}