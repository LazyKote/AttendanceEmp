using System;
using System.Data;
using System.Diagnostics.Contracts;
using System.Windows;
using Google.Protobuf.WellKnownTypes;
using System.Xml.Linq;
using MySql.Data.MySqlClient;
using MySqlX.XDevAPI.Common;
using Mysqlx.Cursor;
using OfficeOpenXml;
using System.IO;

namespace AttendanceEmp
{
    public partial class MainWindow : Window
    {
        private string connectionString;

        public MainWindow()
        {
            InitializeComponent(); // Инициализация XAML главного окна
            connectionString = "Server=localhost;Database=empattend;Uid=root;Pwd=Kate+Kate19;"; // Строка соединения с бд
        }   

        // Метод, загружающий данные из бд в DataGrid
        private void LoadEmployeeData()
        {
            string sql = "SELECT * FROM employees";
            DataTable employeeData = new DataTable();

            // Установка соединения с бд
            try
            {
                using (MySqlConnection connection = new MySqlConnection(connectionString))
                using (MySqlCommand command = new MySqlCommand(sql, connection))// Создание дата адаптера
                using (MySqlDataAdapter adapter = new MySqlDataAdapter(command))
                {
                    connection.Open();
                    adapter.Fill(employeeData);
                }

                // Привязка данных, полученных из бд, к DataGrid
                EmployeeDataGrid.ItemsSource = employeeData.DefaultView;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}");
            }
        }
        // Нажатие кнопки запускает метод, загружающий данные из бд
        private void DataRefresh_Click(object sender, RoutedEventArgs e)
        {           
            LoadEmployeeData();
        }
        // Метод для поиска сотрудника по ФИО или должности
        private void SearchEmployeesByAttribute(string attributeName, string attributeValue)
        {
            string sql = $"SELECT * FROM employees WHERE {attributeName} = @AttributeValue";
            DataTable employeeData = new DataTable();

            try
            {
                using (MySqlConnection connection = new MySqlConnection(connectionString))
                using (MySqlCommand command = new MySqlCommand(sql, connection))
                {
                    // Добавление параметра для предотвращения внедрения SQL-кода
                    command.Parameters.AddWithValue("@AttributeValue", attributeValue);

                    connection.Open();
                    using (MySqlDataAdapter adapter = new MySqlDataAdapter(command))
                    {
                        adapter.Fill(employeeData);
                    }
                }

                EmployeeDataGrid.ItemsSource = employeeData.DefaultView;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}");
            }
        }
        private void DBSearchButton_Click(object sender, RoutedEventArgs e)
        {
            // Получение данных для поиска из TextBox
            string attributeValue = DBSearchBox.Text;
            string attributeName = DBSearchBox.Text;
            // Активная кнопка определяет атрибут, по которому проводится поиск
            if ((bool)PositionRadioButton.IsChecked)
            {
                SearchEmployeesByAttribute("Должность", attributeName);
            }
            else if ((bool)FullNameRadioButton.IsChecked)
            {
                SearchEmployeesByAttribute("ФИО", attributeName);
            }
        }
        //Метод для экспорта данных из бд в Excel-файл
        private void ExportToExcelButton_Click(object sender, RoutedEventArgs e)
        {
            string sql = "SELECT * FROM employees";
            DataTable empData = new DataTable();

            using (MySqlConnection connection = new MySqlConnection(connectionString))
            using (MySqlCommand command = new MySqlCommand(sql, connection))
            using (MySqlDataAdapter adapter = new MySqlDataAdapter(command))
            {
                connection.Open();
                adapter.Fill(empData);
            }

            // Создание нового Excel-файла
            using (ExcelPackage package = new ExcelPackage())
            {
                // Добавление нового листа в Excel-файл
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Employees Data");

                // Заполнение листа данными из бд
                worksheet.Cells["A1"].LoadFromDataTable(empData, true);

                // Сохранение Excel-файла
                string filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "EmployeesData.xlsx");
                FileInfo excelFile = new FileInfo(filePath);
                package.SaveAs(excelFile);
            }

            MessageBox.Show("Данные успешно экспортированы в Excel-файл.");
        }
        // открытие нового окна
        private void NewFormButton_Click(object sender, RoutedEventArgs e)
        {
            Window1 window1 = new Window1();
            window1.Show();
            this.Close();
        }
    }
}