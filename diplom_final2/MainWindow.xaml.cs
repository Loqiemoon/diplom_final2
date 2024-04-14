using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Diagnostics;
using System.IO;
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
using Microsoft.Win32;
using OfficeOpenXml;
using static diplom_final2.Window2;

namespace diplom_final2
{
    public partial class MainWindow : Window
    {
        public string filePath;

        public MainWindow()
        {
            InitializeComponent();
            OpenFile();
            LoadDolj();

            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Excel Files|*.xlsx;*.xlsm",
                Title = "Выберите файл с нагрузкой excel"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                filePath = openFileDialog.FileName;
            }
        }

        private void OpenFile() // Открываю файл в корне
        {
            string exePath = System.Reflection.Assembly.GetExecutingAssembly().Location;
            string excelFilePath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(exePath), "Список преподавателей.xlsx");

            if (System.IO.File.Exists(excelFilePath))
            {
                LoadExcelData(excelFilePath);
            }
            else
            {
                MessageBox.Show("Файл Excel не найден в корне программы.");
            }
        }

        public class DataExcel // Название столбцов и их типы дынных
        {
            public int NN { get; set; }
            public string Фамилия { get; set; }
            public string ФИО { get; set; }
            public string Должность { get; set; }
            public string Бюджет { get; set; }
            public string Внебюджет { get; set; }
            public string Федералы { get; set; }
            public string Часы_по_бюджету { get; set; }
            public string Ставка_по_бюджету { get; set; }
            public string Часы_по_внебюджету { get; set; }
            public string Ставка_по_внебюджету { get; set; }
            public string Ставка_итого { get; set; }
            public string Необходимая_ставка { get; set; }

            public void Value(int columnIndex, object value)
            {
                if (value != null)
                {
                    switch (columnIndex)
                    {
                        case 1:
                            int.TryParse(value.ToString(), out int nnValue);
                            NN = nnValue;
                            break;
                        case 2:
                            Фамилия = value.ToString();
                            break;
                        case 3:
                            ФИО = value.ToString();
                            break;
                        case 4:
                            Должность = value.ToString();
                            break;
                        case 5:
                            Бюджет = value.ToString();
                            break;
                        case 6:
                            Внебюджет = value.ToString();
                            break;
                        case 7:
                            Федералы = value.ToString();
                            break;
                        case 8:
                            Часы_по_бюджету = value.ToString();
                            break;
                        case 9:
                            Ставка_по_бюджету = value.ToString();
                            break;
                        case 10:
                            Часы_по_внебюджету = value.ToString();
                            break;
                        case 11:
                            Ставка_по_внебюджету = value.ToString();
                            break;
                        case 12:
                            Ставка_итого = value.ToString();
                            break;
                        case 13:
                            Необходимая_ставка = value.ToString();
                            break;
                        default:
                            break;
                    }
                }
            }
        }

        private void LoadExcelData(string filePath) // Заношу данные в datagrid
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                List<DataExcel> dataList = new List<DataExcel>();
                int rowCount = worksheet.Dimension.Rows;
                int colCount = worksheet.Dimension.Columns;

                for (int row = 2; row <= 27; row++)
                {
                    DataExcel dataItem = new DataExcel();
                    for (int col = 1; col <= 13; col++)
                    {
                        var cellValue = worksheet.Cells[row, col].Value;
                        dataItem.Value(col, cellValue);
                    }
                    dataList.Add(dataItem);
                }
                raspr.ItemsSource = dataList;
            }
        }

        private void SaveData(string filePath) // Сохранение в excel
        {
            try
            {
                List<DataExcel> dataList = raspr.ItemsSource.Cast<DataExcel>().ToList();
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                    for (int i = 0; i < dataList.Count; i++)
                    {
                        worksheet.Cells[i + 2, 2].Value = dataList[i].Фамилия;
                        worksheet.Cells[i + 2, 3].Value = dataList[i].ФИО;
                        worksheet.Cells[i + 2, 4].Value = dataList[i].Должность;
                        worksheet.Cells[i + 2, 5].Value = dataList[i].Бюджет;
                        worksheet.Cells[i + 2, 6].Value = dataList[i].Внебюджет;
                        worksheet.Cells[i + 2, 7].Value = dataList[i].Федералы;
                        worksheet.Cells[i + 2, 8].Value = dataList[i].Часы_по_бюджету;
                        worksheet.Cells[i + 2, 9].Value = dataList[i].Ставка_по_бюджету;
                        worksheet.Cells[i + 2, 10].Value = dataList[i].Часы_по_внебюджету;
                        worksheet.Cells[i + 2, 11].Value = dataList[i].Ставка_по_внебюджету;
                        worksheet.Cells[i + 2, 12].Value = dataList[i].Ставка_итого;
                        worksheet.Cells[i + 2, 13].Value = dataList[i].Необходимая_ставка;
                    }
                    package.Save();
                }

                MessageBox.Show("Данные успешно сохранены в Excel.", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Произошла ошибка при сохранении данных в Excel: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void SaveDataAuto(string filePath) // Сохранение в excel  не заметно для пользователя
        {
            try
            {
                List<DataExcel> dataList = raspr.ItemsSource.Cast<DataExcel>().ToList();
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                    for (int i = 0; i < dataList.Count; i++)
                    {
                        worksheet.Cells[i + 2, 2].Value = dataList[i].Фамилия;
                        worksheet.Cells[i + 2, 3].Value = dataList[i].ФИО;
                        worksheet.Cells[i + 2, 4].Value = dataList[i].Должность;
                        worksheet.Cells[i + 2, 5].Value = dataList[i].Бюджет;
                        worksheet.Cells[i + 2, 6].Value = dataList[i].Внебюджет;
                        worksheet.Cells[i + 2, 7].Value = dataList[i].Федералы;
                        worksheet.Cells[i + 2, 8].Value = dataList[i].Часы_по_бюджету;
                        worksheet.Cells[i + 2, 9].Value = dataList[i].Ставка_по_бюджету;
                        worksheet.Cells[i + 2, 10].Value = dataList[i].Часы_по_внебюджету;
                        worksheet.Cells[i + 2, 11].Value = dataList[i].Ставка_по_внебюджету;
                        worksheet.Cells[i + 2, 12].Value = dataList[i].Ставка_итого;
                        worksheet.Cells[i + 2, 13].Value = dataList[i].Необходимая_ставка;
                    }

                    package.Save();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Произошла ошибка при сохранении данных в Excel: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)//кнопка сохранения
        {
            SaveData("Список преподавателей.xlsx");
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)//кнопка открытия 2 окна
        {
            SaveDataAuto("Список преподавателей.xlsx");
            Window2 window2 = new Window2(filePath);
            window2.Show();
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)//кнопка чтоб раскидать часы
        {
            Window2 window2 = new Window2(filePath);
            List<Two> BuNameHour = window2.NameHour(window2.Bu);
            ForRasp(BuNameHour, "Бюджет");
            List<Two> NBuNameHour = window2.NameHour(window2.NBu);
            ForRasp(NBuNameHour, "Внебюджет");
            List<Two> FeNameHour = window2.NameHour(window2.Fe);
            ForRasp(FeNameHour, "Федералы");
            raspr.Items.Refresh();
            SumBNF();
            raspr.Items.Refresh();
            SaveDataAuto("Список преподавателей.xlsx");
        }

        private List<string> Doljnost()//список для Должности
        {
            List<string> dolj = new List<string>();
            dolj.Add("Профессор");
            dolj.Add("Доцент");
            dolj.Add("Ст.преподаватель");
            dolj.Add("Преподаватель");
            dolj.Add("Аспирант");

            return dolj;
        }

        private void LoadDolj()//отображаю в выпадающем списке
        {
            List<string> dolj = Doljnost();
            Должности.ItemsSource = dolj;
            Должности.DisplayMemberPath = ".";
        }

        private void ForRasp(List<Two> NameHour, string BNF)//закидываю часы 
        {
            Dictionary<string, float> name = new Dictionary<string, float>();
            foreach (var item in NameHour)
            {
                name[item.Name] = item.Hour;
            }

            foreach (var rowItem in raspr.Items)
            {
                var dataItem = rowItem as DataExcel;
                if (dataItem != null && !string.IsNullOrEmpty(dataItem.ФИО))
                {
                    string fio = dataItem.ФИО;
                    if (name.ContainsKey(fio))
                    {
                        if (BNF == "Бюджет")
                        {
                            dataItem.Бюджет = name[fio].ToString();
                        }
                        if (BNF == "Внебюджет")
                        {
                            dataItem.Внебюджет = name[fio].ToString();
                        }
                        if (BNF == "Федералы")
                        {
                            dataItem.Федералы = name[fio].ToString();
                        }
                    }
                    //raspr.Items.Refresh();
                }
                else
                {
                    break;
                }
            }
        }

        private void SumBNF()//считаю часы и ставки
        {
            foreach (var rowItem in raspr.Items)
            {
                var dataItem = rowItem as DataExcel;
                if (dataItem != null && !string.IsNullOrEmpty(dataItem.ФИО))
                {
                    dataItem.Часы_по_бюджету = (float.Parse(dataItem.Бюджет) + float.Parse(dataItem.Федералы)).ToString("F2");
                    dataItem.Часы_по_внебюджету = dataItem.Внебюджет;
                    if (dataItem.Должность == "Профессор")
                    {
                        dataItem.Ставка_по_бюджету = (float.Parse(dataItem.Часы_по_бюджету) / 600).ToString("F2");
                        dataItem.Ставка_по_внебюджету = (float.Parse(dataItem.Часы_по_внебюджету) / 600).ToString("F2");
                    }
                    else
                    {
                        dataItem.Ставка_по_бюджету = (float.Parse(dataItem.Часы_по_бюджету) / 899).ToString("F2");
                        dataItem.Ставка_по_внебюджету = (float.Parse(dataItem.Часы_по_внебюджету) / 899).ToString("F2");
                    }
                    dataItem.Ставка_итого = (float.Parse(dataItem.Ставка_по_бюджету) + float.Parse(dataItem.Ставка_по_внебюджету)).ToString("F2");
                }
                else
                {
                    break;
                }
            }
            //raspr.Items.Refresh();
        }

        private void Button_Click_3(object sender, RoutedEventArgs e)//открыть excel для печати
        {
            SaveDataAuto("Список преподавателей.xlsx");
            Process.Start("Список преподавателей.xlsx");
        }
    }
}
