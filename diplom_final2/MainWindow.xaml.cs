using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
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

namespace diplom_final2
{
    public partial class MainWindow : Window
    {
        public string FilePath {  get; }

        public MainWindow()
        {
            InitializeComponent();
            OpenFile();
        }
        
        private void OpenFile() // Открываю файл в корне
        {
            string exePath = System.Reflection.Assembly.GetExecutingAssembly().Location;
            string excelFilePath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(exePath), "Распределение нагрузки кафедры по преподавателям.xlsm");

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
            public float? Бюджет { get; set; }
            public float? Внебюджет { get; set; }
            public float? Федералы { get; set; }
            public float? Часы_по_бюджету { get; set; }
            public float? Ставка_по_бюджету { get; set; }
            public float? Часы_по_внебюджету { get; set; }
            public float? Ставка_по_внебюджету { get; set; }
            public float? Ставка_итого { get; set; }
            public float? Необходимая_ставка { get; set; }

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
                            Бюджет = ConvertToFloat(value);
                            break;
                        case 6:
                            Внебюджет = ConvertToFloat(value);
                            break;
                        case 7:
                            Федералы = ConvertToFloat(value);
                            break;
                        case 8:
                            Часы_по_бюджету = ConvertToFloat(value);
                            break;
                        case 9:
                            Ставка_по_бюджету = ConvertToFloat(value);
                            break;
                        case 10:
                            Часы_по_внебюджету = ConvertToFloat(value);
                            break;
                        case 11:
                            Ставка_по_внебюджету = ConvertToFloat(value);
                            break;
                        case 12:
                            Ставка_итого = ConvertToFloat(value);
                            break;
                        case 13:
                            Необходимая_ставка = ConvertToFloat(value);
                            break;
                        default:
                            break;
                    }
                }
            }

            private float? ConvertToFloat(object value)
            {
                if (value != null && float.TryParse(value.ToString(), out float floatValue))
                {
                    return floatValue;
                }
                return null;
            }
        }

        private void LoadExcelData(string filePath) // Заношу данные в datagrid
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                // Загрузка данных в список объектов
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
                        // Добавьте код для остальных свойств
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

        private void Button_Click(object sender, RoutedEventArgs e)//кнопка сохранения
        {
            SaveData("Распределение нагрузки кафедры по преподавателям.xlsm");
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)//кнопка открытия 2 окна
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Excel Files|*.xlsx;*.xlsm",
                Title = "Выберите файл с нагрузкой excel"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                string filePath = openFileDialog.FileName;
                Window2 window2 = new Window2(filePath);
                window2.Show();
            }
        }
    }
}
