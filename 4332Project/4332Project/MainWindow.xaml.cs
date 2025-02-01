using Microsoft.Win32;
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
using Excel = Microsoft.Office.Interop.Excel;

namespace _4332Project
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

        private void SidorovButton_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Имя: Игнат\nВозраст: 18 лет");
        }

        private void ExpBtn_Click(object sender, RoutedEventArgs e)
        {
            List<Clients> allCrid1;
            List<Clients> allCrid2;
            List<Clients> allCrid3;

            using (Lab2Entities lab2Entities = new Lab2Entities())
            {
                var today = DateTime.Today;
                var allClients = lab2Entities.Clients.ToList();

                allCrid1 = allClients
                    .Where(c => c.BirthDate.HasValue && GetAge(c.BirthDate.Value, today) >= 20 && GetAge(c.BirthDate.Value, today) <= 29)
                    .ToList();

                allCrid2 = allClients
                    .Where(c => c.BirthDate.HasValue && GetAge(c.BirthDate.Value, today) >= 30 && GetAge(c.BirthDate.Value, today) <= 39)
                    .ToList();

                allCrid3 = allClients
                    .Where(c => c.BirthDate.HasValue && GetAge(c.BirthDate.Value, today) >= 40)
                    .ToList();

            }

            SaveFileDialog sfd = new SaveFileDialog()
            {
                Filter = "Excel Files|*.xlsx",
                Title = "Сохранить файл Excel",
                FileName = "UsersByAge.xlsx"
            };

            if (sfd.ShowDialog() == true)
            {
                ExportToExcel(sfd.FileName, allCrid1, allCrid2, allCrid3);
                MessageBox.Show("Файл успешно сохранен!", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private int GetAge(DateTime birthDate, DateTime today)
        {
            int age = today.Year - birthDate.Year;
            if (birthDate > today.AddYears(-age)) age--;
            return age;
        }

        private void ExportToExcel(string filePath, List<Clients> crid1, List<Clients> crid2, List<Clients> crid3)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = excelApp.Workbooks.Add();

            CreateSheet(workbook, "20-29 лет", crid1);
            CreateSheet(workbook, "30-39 лет", crid2);
            CreateSheet(workbook, "40+ лет", crid3);

            workbook.SaveAs(filePath);
            workbook.Close();
            excelApp.Quit();

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        private void CreateSheet(Excel.Workbook workbook, string sheetName, List<Clients> clients)
        {
            Excel.Worksheet worksheet = workbook.Sheets.Add();
            worksheet.Name = sheetName;

            worksheet.Cells[1, 1] = "ID";
            worksheet.Cells[1, 2] = "ФИО";
            worksheet.Cells[1, 3] = "Дата рождения";

            int row = 2;
            foreach (var client in clients)
            {
                worksheet.Cells[row, 1] = client.Client_ID;
                worksheet.Cells[row, 2] = client.FIO;

                // Проверяем на null, чтобы безопасно использовать ToShortDateString()
                if (client.BirthDate.HasValue)
                {
                    worksheet.Cells[row, 3] = client.BirthDate.Value.ToShortDateString();
                }
                else
                {
                    worksheet.Cells[row, 3] = "Не указана";
                }

                row++;
            }

            worksheet.Columns.AutoFit();
        }


        private void ImpBtn_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.xls; *.xlsx",
                Filter = "файл Excel (Spisok.xlsx)|*.xlsx",
                Title = "Выберите файл базы данных"
            };

            if (!(ofd.ShowDialog() == true)) { return; }

            string[,] list;
            Excel.Application ObjWorkExcel = new Excel.Application();
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(ofd.FileName);
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];

            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);

            int _rows = (int)lastCell.Row;
            int _columns = (int)lastCell.Column;
            list = new string[_rows, _columns];

            for (int j = 0; j < _columns; j++)
            {
                for (int i = 0; i < _rows; i++)
                {
                    list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text;
                }
            }

            ObjWorkBook.Close(false, Type.Missing, Type.Missing);
            ObjWorkExcel.Quit();
            GC.Collect();

            using (Lab2Entities lab2Entities = new Lab2Entities())
            {
                for (int i = 0; i < _rows; i++)
                {
                    // Parse birthdate if present, else set to null
                    DateTime? birthDate = null;
                    if (DateTime.TryParse(list[i, 2], out DateTime parsedDate)) // Assuming BirthDate is in column 3 (index 2)
                    {
                        birthDate = parsedDate;
                    }

                    lab2Entities.Clients.Add(new Clients()
                    {
                        FIO = list[i, 0],
                        Email = list[i, 8],
                        BirthDate = birthDate // Assign parsed birth date
                    });
                    lab2Entities.SaveChanges();
                }
            }

            MessageBox.Show("Успешно!");
        }

    }
}
