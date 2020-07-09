using System;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using ClosedXML.Excel;
using FSTEC_Analytics.Models;
using Xceed.Words.NET;

namespace FSTEC_Analytics.Logic
{

    internal class Parser
    {
        public event Action<bool> IsDoneSuccess;

        private string fileName;
        private DateTime dateFrom;
        private DateTime dateTo;
        private bool buttonKselected;
        private string outputFileName;

        private VulnerabilityModel vulnerability = new VulnerabilityModel();

        public Parser(string fileName)
        {
            this.fileName = fileName;
        }

        public async void Start(DateTime dateFrom, DateTime dateTo, bool buttonKselected, string outputFileName)
        {
            this.dateFrom = dateFrom;
            this.dateTo = dateTo;
            this.buttonKselected = buttonKselected;
            this.outputFileName = outputFileName;

            await Task.Run(() => ParseExcel());
        }

        private void ParseExcel()
        {
            try
            {
                using (var workbook = new XLWorkbook(fileName, XLEventTracking.Disabled))
                {
                    var worksheet = workbook.Worksheet(1);

                    var c = 0; // Количество элементов подходящие под гугл хром

                    foreach (var row in worksheet.RowsUsed().Skip(2))
                    {
                        if (row.Cell(5).Value.ToString() == "Google Chrome" /*&& (DateTime)row.Cell(10).Value <= DateTime*/)
                        {
                            c++;

                            VulnerabilityObject item = new VulnerabilityObject
                            {
                                ID = row.Cell(1).Value.ToString(),
                                Title = row.Cell(2).Value.ToString(),
                                Description = row.Cell(3).Value.ToString(),
                                VendorPO = row.Cell(4).Value.ToString(),
                                NamePO = row.Cell(5).Value.ToString(),
                                VersionPO = row.Cell(6).Value.ToString(),
                                TypePO = row.Cell(7).Value.ToString(),
                                OS = row.Cell(8).Value.ToString(),
                                Class = row.Cell(9).Value.ToString(),
                                Date = (DateTime)row.Cell(10).Value,
                                CVSS2 = row.Cell(11).Value.ToString(),
                                CVSS3 = row.Cell(12).Value.ToString(),
                                DangerLevel = row.Cell(13).Value.ToString(),
                                PossibleCorrective = row.Cell(14).Value.ToString(),
                                Status = row.Cell(15).Value.ToString(),
                                Exploit = row.Cell(16).Value.ToString(),
                                Information = row.Cell(17).Value.ToString(),
                                URL = row.Cell(18).Value.ToString(),
                                OtherIdentifiers = row.Cell(19).Value.ToString(),
                                Etc = row.Cell(20).Value.ToString()
                            };

                            vulnerability.AddItem(item);
                        }
                    }

                    if (buttonKselected)
                    {
                        PutOnFile();
                    }
                    else
                    {
                        PutOnFileFull();
                    }
                }
            }
            catch (ArgumentException)
            {
                MessageBox.Show("Программа поддерживает анализ файлов только данных банка уязвимостей ФСТЭК в формате\".xlsx\"\nПожалуйста, повторите попытку!", "ОШИБКА", MessageBoxButton.OK, MessageBoxImage.Error);
                IsDoneSuccess?.Invoke(false);
            }
            catch (InvalidCastException)
            {
                MessageBox.Show("Программа поддерживает анализ файлов только данных банка уязвимостей ФСТЭК в формате\".xlsx\"\nПожалуйста, повторите попытку!", "ОШИБКА", MessageBoxButton.OK, MessageBoxImage.Error);
                IsDoneSuccess?.Invoke(false);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "ОШИБКА", MessageBoxButton.OK, MessageBoxImage.Error);
                IsDoneSuccess?.Invoke(false);
            }
        }

        private void PutOnFileFull()
        {
            using (var document = DocX.Create(outputFileName))
            {
                try
                {
                    var paragrahp = document.InsertParagraph();

                    VulnerabilityModel filter = vulnerability.Report(dateFrom, dateTo);

                    for (int i = 0; i < filter.Count; i++)
                    {
                        var item = filter.GetItem(i);
                        paragrahp.Append($"{item.Date:D} | {item.ID} | {item.Title} | {item.NamePO}\n\n");
                    }

                    document.Save();
                    IsDoneSuccess?.Invoke(true);
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "ОШИБКА", MessageBoxButton.OK, MessageBoxImage.Error);
                    IsDoneSuccess?.Invoke(false);
                }
            }
        }

        private void PutOnFile()
        {
            using (var document = DocX.Create(outputFileName))
            {
                try
                {
                    var paragrahp = document.InsertParagraph();

                    VulnerabilityModel filter = vulnerability.Report(dateFrom, dateTo);

                    paragrahp.Append("Распределение количества уязвимостей по годам:\n");
                    foreach (var item in vulnerability.ObjectListYear)
                    {
                        paragrahp.Append($"{item.Year} г. - {item.Count}\n");
                    }
                    paragrahp.Append("\n\n\n");

                    paragrahp.Append($"Уровни опасности за {dateFrom:D} по {dateTo:D}\n"); ;
                    paragrahp.Append($"Критический уровень: {vulnerability.Critical} уязвимостей\n" +
                                     $"Высокий уровень: {vulnerability.High} уязвимостей\n" +
                                     $"Средний уровень: {vulnerability.Middle} уязвимостей\n" +
                                     $"Высокий уровень: {vulnerability.Low} уязвимостей");

                    document.Save();
                    IsDoneSuccess?.Invoke(true);
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "ОШИБКА", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }
    }
}