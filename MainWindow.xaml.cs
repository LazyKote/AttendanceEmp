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
            connectionString = "Server=localhost;Database=empattend;Uid=root;Pwd=Kate+Kate19;"; // Установка соединения с бд
        }   

        // Method to load employee data from the database and populate the DataGrid
        private void LoadEmployeeData()
        {
            string sql = "SELECT * FROM employees";// SQL query to select all columns from the employees table
            DataTable employeeData = new DataTable();

            // Establish a connection to the MySQL database
            try
            {
                using (MySqlConnection connection = new MySqlConnection(connectionString))// Create a command to execute the SQL query
                using (MySqlCommand command = new MySqlCommand(sql, connection))// Create a data adapter to fill the DataTable with the query results
                using (MySqlDataAdapter adapter = new MySqlDataAdapter(command))// Execute the query and fill the DataTable
                {
                    connection.Open();
                    adapter.Fill(employeeData);
                }

                // Bind the DataTable to the DataGrid
                EmployeeDataGrid.ItemsSource = employeeData.DefaultView;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}");
            }
        }

        private void DataRefresh_Click(object sender, RoutedEventArgs e)
        {
            // This button click event can be used to refresh the employee data
            LoadEmployeeData();
        }
        // Method to search for employees by position
        private void SearchEmployeesByAttribute(string attributeName, string attributeValue)
        {
            // Construct the SQL query dynamically based on the attribute name and value
            string sql = $"SELECT * FROM employees WHERE {attributeName} = @AttributeValue";
            DataTable employeeData = new DataTable();

            try
            {
                using (MySqlConnection connection = new MySqlConnection(connectionString))
                using (MySqlCommand command = new MySqlCommand(sql, connection))
                {
                    // Add parameter to prevent SQL injection
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
                MessageBox.Show($"An error occurred: {ex.Message}");
            }
        }
        private void DBSearchButton_Click(object sender, RoutedEventArgs e)
        {
            // Get the search term from the TextBox
            string attributeValue = DBSearchBox.Text;
            string attributeName = DBSearchBox.Text;
            // Determine which RadioButton is selected and search by the corresponding attribute
            if ((bool)PositionRadioButton.IsChecked)
            {
                SearchEmployeesByAttribute("Должность", attributeName);
            }
            else if ((bool)FullNameRadioButton.IsChecked)
            {
                SearchEmployeesByAttribute("ФИО", attributeName);
            }
        }
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

                // Заполнение листа данными из DataTable
                worksheet.Cells["A1"].LoadFromDataTable(empData, true);

                // Сохранение Excel-файла
                string filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "EmployeesData.xlsx");
                FileInfo excelFile = new FileInfo(filePath);
                package.SaveAs(excelFile);
            }

            MessageBox.Show("Данные успешно экспортированы в Excel-файл.");
        }

        private void NewFormButton_Click(object sender, RoutedEventArgs e)
        {
            Window1 window1 = new Window1();
            window1.Show();
        }
    }
}