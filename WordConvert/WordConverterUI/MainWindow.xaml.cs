using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Forms;
using WordConvert;

namespace WordConverterUI
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void browse_Click(object sender, RoutedEventArgs e)
        {
            // Create OpenFileDialog
            var dlg = new System.Windows.Forms.FolderBrowserDialog();

            // Display OpenFileDialog by calling ShowDialog method
            var result = dlg.ShowDialog();

            if(result == System.Windows.Forms.DialogResult.OK)
            {
                FileNameTextBox.Text = dlg.SelectedPath;
            }
        }

        private void submit_Click(object sender, RoutedEventArgs e)
        {
            var newFolder = FileNameTextBox.Text + "_converted";
            var result = WordConvert.WordConverter.ConvertPath(FileNameTextBox.Text, newFolder);
            if (result != 0)
            {
                statusLabel.Content = "Конвертация прошла успешно: " + newFolder + ", " + Convert.ToString(result) + " файлов сконвертировано.";
            }
            else
            {
                statusLabel.Content = "Ошибка! Не найдено файлов для конвертации по пути: " + FileNameTextBox.Text;
            }
        }
    }
}
