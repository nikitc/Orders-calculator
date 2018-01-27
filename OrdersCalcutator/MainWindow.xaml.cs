using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Win32;

namespace OrdersCalcutator
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow
    {
        public MainWindow()
        {
            InitializeComponent();
            ResizeMode = ResizeMode.NoResize;
        }

        private void ChooseFile_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog
            {
                Filter = "Таблицы(*.xls;*.xlsx)|*.xls;*.xlsx" + "|Все файлы (*.*)|*.* ",
                CheckFileExists = true,
                Multiselect = true
            };

            if (openFileDialog.ShowDialog() == true)
            {
                if (!string.IsNullOrEmpty(FilePath.Text))
                    FilePath.Text += "\n"; 
                FilePath.Text += string.Join("\n", openFileDialog.FileNames);
            }
        }


        private async void CalcResult_Click(object sender, RoutedEventArgs e)
        {
            if (IsHasError())
            {
                ResultText.Text = GetErrorText();
                ResultText.Foreground = System.Windows.Media.Brushes.Red;
                return;
            }

            ResultText.Text = "Процесс...";
            ResultText.Foreground = System.Windows.Media.Brushes.Black;
            var files = FilePath.Text.Split('\n').ToArray();
            var startDate = (DateTime)StartDate.SelectedDate;
            var finishDate = (DateTime)FinishDate.SelectedDate;

            await Task.Factory.StartNew(() => OrdersCalculator.CalculateOrders(files, startDate, finishDate));

            ResultText.Text = $"Успех! Результат в текущей папке, в файле Result_Calc_orders.xls";
            ResultText.Foreground = System.Windows.Media.Brushes.Green;
        }

        private bool IsHasError()
        {
            return FilePath.Text == "" || FilePath.Text.Split('\n').Any(path => !File.Exists(path)) 
                || StartDate.SelectedDate == null || FinishDate.SelectedDate == null 
                || StartDate.SelectedDate > FinishDate.SelectedDate;
        }

        private string GetErrorText()
        {
            var error = "Ошибка:\n";
            if (FilePath.Text == "")
                error += "Отсутствует путь до файла\n";
            else
            {
                error = FilePath.Text
                    .Split('\n')
                    .Where(path => !File.Exists(path))
                    .Aggregate(error, (current, path) => current + $"Неверный путь файла {path}\n");
            }

            if (StartDate.SelectedDate == null)
                error += "Отсутствует стартовая дата\n";

            if (FinishDate.SelectedDate == null)
                error += "Отсутствует конечная дата\n";

            if (StartDate.SelectedDate != null && FinishDate.SelectedDate != null
                && StartDate.SelectedDate > FinishDate.SelectedDate)
                error += "Конечная дата должна быть больше стартовой";

            return error;
        }
    }
}
