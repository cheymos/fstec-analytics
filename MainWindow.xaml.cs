using System;
using System.IO;
using System.Net;
using System.Windows;
using System.Windows.Input;
using FSTEC_Analytics.Logic;
using Microsoft.Win32;

namespace FSTEC_Analytics
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        string fileName;
        bool buttonKselected = false;
        string outputFileName = "report_FSTEC.docx";
        string dowloadLink = @"https://bdu.fstec.ru/files/documents/vullist.xlsx";

        public MainWindow()
        {
            InitializeComponent();
        }

        private void GridFieldDrop_DragEnter(object sender, DragEventArgs e)
        {
            DragText.Text = "Можете отпускать!";
        }

        private void GridFieldDrop_DragLeave(object sender, DragEventArgs e)
        {
            DragText.Text = "Сюда можно бросить файл!";

        }

        private void GridFieldDrop_Drop(object sender, DragEventArgs e)
        {
            var files = (string[])e.Data.GetData(DataFormats.FileDrop);
            fileName = files[0].ToString();

            DragText.Text = fileName;
        }

        private void GridFieldDrop_MouseDown(object sender, MouseButtonEventArgs e) // Вызов OpenFileDialog
        {
            var openFileDialog = new OpenFileDialog
            {
                Filter = "Excel Worksheets|*.xlsx",
                Title = "Выберите таблицу Excel"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                if (openFileDialog.CheckFileExists)
                {
                    fileName = openFileDialog.FileName;
                    DragText.Text = fileName;
                }
            }
        }



        private void Button_Click(object sender, RoutedEventArgs e)
        {
            if (DateTime.TryParse(DateFrom.Text, out DateTime dateFrom) && DateTime.TryParse(DateTo.Text, out DateTime dateTo))
            {
                if (fileName != null)
                {
                    if (File.Exists(fileName))
                    {
                        var result = MessageBox.Show($"Готовый отчёт уже существует и будет перезаписан\nПродолжить выполнение?", "Информация", MessageBoxButton.OKCancel, MessageBoxImage.Asterisk);
                        if (result == MessageBoxResult.OK)
                        {
                            var parser = new Parser(fileName);

                            parser.Start(dateFrom, dateTo, buttonKselected, outputFileName);
                            parser.IsDoneSuccess += Parser_IsDoneSuccess;

                            Window.IsEnabled = false;
                        }
                        else
                        {
                            return;
                        }
                    }
                }
                else
                {
                    MessageBox.Show("Файл не выбран.\nВыберите файл и повторите попытку!", "ОШИБКА", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            else
            {
                MessageBox.Show("Некорректно введена дата!", "ОШИБКА", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void Parser_IsDoneSuccess(bool obj)
        {
            Dispatcher.Invoke(new Action(delegate () { Window.IsEnabled = true; DragText.Text = "Сюда можно бросить файл!"; }));
            if (obj == true)
                MessageBox.Show($"Отчёт [{outputFileName}] успешно сформирован в директории рядом с программой", "Информация", MessageBoxButton.OK, MessageBoxImage.Information, MessageBoxResult.OK);
        }

        private void Button_P_Click(object sender, RoutedEventArgs e)
        {
            buttonKselected = false;

            if (Button_P.IsEnabled)
            {
                Button_P.IsEnabled = false;
                Button_K.IsEnabled = true;
            }
        }

        private void Button_K_Click(object sender, RoutedEventArgs e)
        {
            buttonKselected = true;

            if (Button_K.IsEnabled)
            {
                Button_K.IsEnabled = false;
                Button_P.IsEnabled = true;
            }
        }

        private void DownloadFile()
        {
            fileName = "FSTEC_file.xlsx";
            if (File.Exists($"{fileName}"))
            {
                var result = MessageBox.Show($"Файл данных банка уязвимостей ФСТЭК уже был загружен раннее\nСкачать заново?", "Информация", MessageBoxButton.YesNo, MessageBoxImage.Asterisk);


                if (result == MessageBoxResult.Yes)
                {
                    File.Delete($"{fileName}");
                }
                else if (result == MessageBoxResult.No)
                {
                    fileName = $"{fileName}";
                    DragText.Text = fileName;
                    return;
                }
            }
            WebClient webClient = new WebClient();
            webClient.DownloadFileAsync(new Uri(dowloadLink), ($"{fileName}"));
            webClient.DownloadFileCompleted += WebClient_DownloadFileCompleted;
            Window.IsEnabled = false;
        }

        private void WebClient_DownloadFileCompleted(object sender, System.ComponentModel.AsyncCompletedEventArgs e)
        {
            MessageBox.Show($"Файл загружен!", "Информация", MessageBoxButton.OK, MessageBoxImage.Information, MessageBoxResult.OK);

            Parser_IsDoneSuccess(false);
            DragText.Text = fileName;
        }

        private void TextBlock_MouseDown(object sender, MouseButtonEventArgs e)
        {
            DownloadFile();
        }
    }
}