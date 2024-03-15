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
    public class ExcelDataItem
    {
        public string NN { get; set; }
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
        public string Необхадимая_ставка { get; set; }
    }

    public partial class MainWindow : Window
    {
        private ObservableCollection<ExcelDataItem> excelDataCollection;

        public MainWindow()
        {
            InitializeComponent();
            excelDataCollection = new ObservableCollection<ExcelDataItem>();
            raspr.ItemsSource = excelDataCollection;
            OpenFile();
        }

        private void OpenFile()
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

        private void LoadExcelData(string filePath)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            excelDataCollection.Clear();

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.Rows;
                int colCount = worksheet.Dimension.Columns;

                for (int row = 2; row <= rowCount; row++)
                {
                    ExcelDataItem item = new ExcelDataItem();
                    for (int col = 1; col <= colCount; col++)
                    {
                        var cellValue = worksheet.Cells[row, col].Value;

                        switch (col)
                        {
                            case 1:
                                item.NN = cellValue?.ToString();
                                break;
                            case 2:
                                item.Фамилия = cellValue?.ToString();
                                break;
                            case 3:
                                item.ФИО = cellValue?.ToString();
                                break;
                            case 4:
                                item.Должность = cellValue?.ToString();
                                break;
                            case 5:
                                item.Бюджет = cellValue?.ToString();
                                break;
                            case 6:
                                item.Внебюджет = cellValue?.ToString();
                                break;
                            case 7:
                                item.Федералы = cellValue?.ToString();
                                break;
                            case 8:
                                item.Часы_по_бюджету = cellValue?.ToString();
                                break;
                            case 9:
                                item.Ставка_по_бюджету = cellValue?.ToString();
                                break;
                            case 10:
                                item.Часы_по_внебюджету = cellValue?.ToString();
                                break;
                            case 11:
                                item.Ставка_по_внебюджету = cellValue?.ToString();
                                break;
                            case 12:
                                item.Ставка_итого = cellValue?.ToString();
                                break;
                            case 13:
                                item.Необхадимая_ставка = cellValue?.ToString();
                                break;
                        }
                    }
                    excelDataCollection.Add(item);
                }
            }
        }


        private void Button_Click(object sender, RoutedEventArgs e)
        {
            
        }
    }
}
