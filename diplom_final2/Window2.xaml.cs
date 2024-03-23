using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.IO.Packaging;
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
using System.Windows.Shapes;
using static diplom_final2.MainWindow;

namespace diplom_final2
{
    public partial class Window2 : Window
    {
        private string _filePath;
        public Window2(string filePath)
        {
            InitializeComponent();
            _filePath = filePath;
            LoadBudgetData(_filePath);
            LoadNonBudgetData(_filePath);
            LoadFederalData(_filePath);
        }

        public class DataExcel_2
        {
            public string Шифр { get; set; }
            public string Преподаватель { get; set; }
            public string Дисциплина { get; set; }
            public string Группа { get; set; }
            public int? Курс { get; set; }
            public int? Семестр { get; set; }
            public int? Кол_чел { get; set; }
            public float? Лекции_1 { get; set; }
            public float? Практические_занятия_1 { get; set; }
            public float? Курсовые_работы_начитка_1 { get; set; }
            public float? Проверка_курсовой_работы_1 { get; set; }
            public float? Контрольные_работы_1 { get; set; }
            public float? РГР_1 { get; set; }
            public float? Текущие_консультации_1 { get; set; }
            public float? Зачет_1 { get; set; }
            public float? Зачет_оценка_1 { get; set; }
            public float? Экзамен_консультация_1 { get; set; }
            public float? Всего_за_семестр_1 { get; set; }
            public float? Лекции_2 { get; set; }
            public float? Практические_занятия_2 { get; set; }
            public float? Курсовые_работы_начитка_2 { get; set; }
            public float? Проверка_курсовой_работы_2 { get; set; }
            public float? Контрольные_работы_2 { get; set; }
            public float? РГР_2 { get; set; }
            public float? Текущие_консультации_2 { get; set; }
            public float? Зачет_2 { get; set; }
            public float? Зачет_оценка_2 { get; set; }
            public float? Экзамен_консультация_2 { get; set; }
            public float? Всего_за_семестр_2 { get; set; }
            public float? Всего_за_год { get; set; }

            public void Value_2(int columnIndex, object value)
            {
                switch (columnIndex)
                {
                    case 1:
                        Шифр = Convert.ToString(value);
                        break;
                    case 2:
                        Преподаватель = Convert.ToString(value);
                        break;
                    case 3:
                        Дисциплина = Convert.ToString(value);
                        break;
                    case 4:
                        Группа = Convert.ToString(value);
                        break;
                    case 5:
                        Курс = ConvertToInt(value);
                        break;
                    case 6:
                        Семестр = ConvertToInt(value);
                        break;
                    case 7:
                        Кол_чел = ConvertToInt(value);
                        break;
                    case 8:
                        Лекции_1 = ConvertToFloat(value);
                        break;
                    case 9:
                        Практические_занятия_1 = ConvertToFloat(value);
                        break;
                    case 10:
                        Курсовые_работы_начитка_1 = ConvertToFloat(value);
                        break;
                    case 11:
                        Проверка_курсовой_работы_1 = ConvertToFloat(value);
                        break;
                    case 12:
                        Контрольные_работы_1 = ConvertToFloat(value);
                        break;
                    case 13:
                        РГР_1 = ConvertToFloat(value);
                        break;
                    case 14:
                        Текущие_консультации_1 = ConvertToFloat(value);
                        break;
                    case 15:
                        Зачет_1 = ConvertToFloat(value);
                        break;
                    case 16:
                        Зачет_оценка_1 = ConvertToFloat(value);
                        break;
                    case 17:
                        Экзамен_консультация_1 = ConvertToFloat(value);
                        break;
                    case 18:
                        Всего_за_семестр_1 = ConvertToFloat(value);
                        break;
                    case 19:
                        Лекции_2 = ConvertToFloat(value);
                        break;
                    case 20:
                        Практические_занятия_2 = ConvertToFloat(value);
                        break;
                    case 21:
                        Курсовые_работы_начитка_2 = ConvertToFloat(value);
                        break;
                    case 22:
                        Проверка_курсовой_работы_2 = ConvertToFloat(value);
                        break;
                    case 23:
                        Контрольные_работы_2 = ConvertToFloat(value);
                        break;
                    case 24:
                        РГР_2 = ConvertToFloat(value);
                        break;
                    case 25:
                        Текущие_консультации_2 = ConvertToFloat(value);
                        break;
                    case 26:
                        Зачет_2 = ConvertToFloat(value);
                        break;
                    case 27:
                        Зачет_оценка_2 = ConvertToFloat(value);
                        break;
                    case 28:
                        Экзамен_консультация_2 = ConvertToFloat(value);
                        break;
                    case 29:
                        Всего_за_семестр_2 = ConvertToFloat(value);
                        break;
                    case 30:
                        Всего_за_год = ConvertToFloat(value);
                        break;
                    default:
                        break;
                }
            }
            private int? ConvertToInt(object value)
            {
                if (value != null && int.TryParse(value.ToString(), out int intValue))
                {
                    return intValue;
                }
                return null;
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

        private async void LoadBudgetData(string filePath)
        {
            await Task.Run(() => LoadExcelData(filePath, "Бюджет", Bu));
            UpdateUI();
        }

        private async void LoadNonBudgetData(string filePath)
        {
            await Task.Run(() => LoadExcelData(filePath, "Внебюджет", NBu));
            UpdateUI();
        }

        private async void LoadFederalData(string filePath)
        {
            await Task.Run(() => LoadExcelData(filePath, "Федералы", Fe));
            UpdateUI();
        }

        private Task LoadExcelData(string filePath, string sheetName, DataGrid dataGrid)
        {
            try
            {
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[sheetName];

                    List<DataExcel_2> dataList = new List<DataExcel_2>();
                    int rowCount = worksheet.Dimension.Rows;
                    int colCount = worksheet.Dimension.Columns;

                    for (int row = 4; row <= 10; row++)
                    {
                        DataExcel_2 dataItem = new DataExcel_2();
                        for (int col = 1; col <= 31; col++)
                        {
                            var cellValue = worksheet.Cells[row, col].Value;
                            dataItem.Value_2(col, cellValue);
                        }
                        dataList.Add(dataItem);
                    }
                    Dispatcher.Invoke(() =>
                    {
                        Console.WriteLine($"Loaded {dataList.Count} items.");
                        dataGrid.ItemsSource = dataList;
                    });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Произошла ошибка при загрузке данных из Excel: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }

            return Task.CompletedTask;
        }

        private void UpdateUI()
        {
          
        }
    }
}

