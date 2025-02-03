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
using System.Text.Json;
using System.IO;
using Newtonsoft.Json.Linq;
using System.Xml;
using System.Text.Json.Serialization;
using System.Globalization;





namespace _4332Project
{
    

    //кастомный конвертатор даты для импоорта json
    public class DateTimeConverter : JsonConverter<DateTime?>
    {
        private const string DateFormat = "dd.MM.yyyy";

        public override DateTime? Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
        {
            if (reader.TokenType == JsonTokenType.Null)
            {
                return null;
            }

            if (reader.TokenType == JsonTokenType.String)
            {
                string dateString = reader.GetString();
                if (DateTime.TryParseExact(dateString, DateFormat, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime parsedDate))
                {
                    return parsedDate;
                }
            }

            return null; 
        }

        public override void Write(Utf8JsonWriter writer, DateTime? value, JsonSerializerOptions options)
        {
            if (value.HasValue)
            {
                writer.WriteStringValue(value.Value.ToString(DateFormat));
            }
            else
            {
                writer.WriteNullValue();
            }
        }
    }

    
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
                    DateTime? birthDate = null;
                    if (DateTime.TryParse(list[i, 2], out DateTime parsedDate))
                    {
                        birthDate = parsedDate;
                    }

                    lab2Entities.Clients.Add(new Clients()
                    {
                        FIO = list[i, 0],
                        Email = list[i, 8],
                        BirthDate = birthDate
                    });
                    lab2Entities.SaveChanges();
                }
            }

            MessageBox.Show("Успешно!");
        }

        static string FormatJsonFile(string json, string filePath)
        {
            try
            {
                
                var jsonArray = JArray.Parse(json);

                
                string formattedJson = jsonArray.ToString(Newtonsoft.Json.Formatting.Indented);

                
                File.WriteAllText(filePath, formattedJson);
                MessageBox.Show(formattedJson, "Содержимое JSON", MessageBoxButton.OK, MessageBoxImage.Information);

                return formattedJson;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при обработке файла: {ex.Message}");
                return json;
            }
        }

        private void ImpJsonBtn_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.json;",
                Filter = "Json файлы (*.json)|*.json",
                Title = "Выберите JSON файл базы данных"
            };

            if (ofd.ShowDialog() != true) { return; }

            string json = File.ReadAllText(ofd.FileName);

            // Десериализация JSON в список объектов
            var importedClients = JsonSerializer.Deserialize<List<Clients>>(json, new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true,
                Converters = { new DateTimeConverter() } 
            });

            if (importedClients != null)
            {
                using (Lab2Entities lab2Entities = new Lab2Entities())
                {
                    lab2Entities.Clients.AddRange(importedClients);
                    lab2Entities.SaveChanges();
                }

                MessageBox.Show("Данные успешно загружены!", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            else
            {
                MessageBox.Show("Ошибка при загрузке файла!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }


        private void ExpJsonInWordBtn_Click(object sender, RoutedEventArgs e)
        {
            // Открытие диалога выбора файла для импорта JSON
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.json;",
                Filter = "Json файлы (*.json)|*.json",
                Title = "Выберите JSON файл базы данных"
            };

            if (ofd.ShowDialog() != true) { return; }

            // Чтение содержимого JSON файла
            string json = File.ReadAllText(ofd.FileName);

            // Десериализация JSON в список объектов
            var importedClients = JsonSerializer.Deserialize<List<Clients>>(json, new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true,
                Converters = { new DateTimeConverter() }
            });

            if (importedClients != null)
            {
                // Группируем клиентов по возрастным категориям
                var category1 = importedClients.Where(c => c.BirthDate.HasValue && GetAge(c.BirthDate.Value, DateTime.Today) >= 20 && GetAge(c.BirthDate.Value, DateTime.Today) <= 29).ToList();
                var category2 = importedClients.Where(c => c.BirthDate.HasValue && GetAge(c.BirthDate.Value, DateTime.Today) >= 30 && GetAge(c.BirthDate.Value, DateTime.Today) <= 39).ToList();
                var category3 = importedClients.Where(c => c.BirthDate.HasValue && GetAge(c.BirthDate.Value, DateTime.Today) >= 40).ToList();

                // Создание нового Word документа
                var wordApp = new Microsoft.Office.Interop.Word.Application();
                var wordDoc = wordApp.Documents.Add();

                // Добавление категории 1 - 20-29 лет
                AddCategoryToWord(wordDoc, "Категория 1 (20-29 лет)", category1);

                // Добавление категории 2 - 30-39 лет
                AddCategoryToWord(wordDoc, "Категория 2 (30-39 лет)", category2);

                // Добавление категории 3 - 40+ лет
                AddCategoryToWord(wordDoc, "Категория 3 (40+ лет)", category3);

                // Сохранение документа
                SaveFileDialog sfd = new SaveFileDialog()
                {
                    Filter = "Word Files|*.docx",
                    Title = "Сохранить файл Word",
                    FileName = "ClientsByAge.docx"
                };

                if (sfd.ShowDialog() == true)
                {
                    wordDoc.SaveAs2(sfd.FileName);
                    wordDoc.Close();
                    wordApp.Quit();

                    MessageBox.Show("Файл успешно сохранен!", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            else
            {
                MessageBox.Show("Ошибка при загрузке файла!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void AddCategoryToWord(Microsoft.Office.Interop.Word.Document wordDoc, string categoryTitle, List<Clients> clients)
        {
            // Вставляем заголовок категории
            var paragraph = wordDoc.Paragraphs.Add();
            paragraph.Range.Text = categoryTitle;
            paragraph.Range.set_Style("Заголовок 1"); // Используем строковое имя стиля для заголовка
            paragraph.Range.InsertParagraphAfter();

            // Вставляем список клиентов в эту категорию
            foreach (var client in clients)
            {
                var clientInfo = $"{client.FIO} - {client.BirthDate?.ToString("dd.MM.yyyy") ?? "Не указана"}";
                paragraph = wordDoc.Paragraphs.Add();
                paragraph.Range.Text = clientInfo;
                paragraph.Range.InsertParagraphAfter();
            }

            // Добавление разрыва страницы после каждой категории
            paragraph = wordDoc.Paragraphs.Add();
            paragraph.Range.InsertBreak(Microsoft.Office.Interop.Word.WdBreakType.wdPageBreak);
        }




    }
}
