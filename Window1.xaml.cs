using MySql.Data.MySqlClient;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
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
using System.Windows.Shapes;

namespace AttendanceEmp
{
    /// <summary>
    /// Логика взаимодействия для Window1.xaml
    /// </summary>
    public partial class Window1 : Window
    {
        double totalHoursWorked = 0;
        private string connectionString;
        public Window1()
        {
            InitializeComponent();
            connectionString = "Server=localhost;Database=empattend;Uid=root;Pwd=Kate+Kate19;"; // Установка соединения с бд
        }
        private void LoadAttendenceData()
        {
            string sql = "SELECT * FROM attendance"; // SQL query to select all columns from the attendance table
            DataTable attendData = new DataTable();

            // Establish a connection to the MySQL database
            try
            {
                using (MySqlConnection connection = new MySqlConnection(connectionString))
                using (MySqlCommand command = new MySqlCommand(sql, connection))
                using (MySqlDataAdapter adapter = new MySqlDataAdapter(command))
                {
                    connection.Open();
                    adapter.Fill(attendData);
                }

                // Bind the DataTable to the DataGrid
                AttendenceData.ItemsSource = attendData.DefaultView;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}");
            }
        }

        private void AttendDataLoad_Click(object sender, RoutedEventArgs e)
        {
            // This button click event can be used to refresh the attendance data
            LoadAttendenceData();
        }

        public double CalculateHoursWorkedForMonth(string employeeName, int month, int year)
        {
            double totalHoursWorked = 0;

            using (MySqlConnection connection = new MySqlConnection(connectionString))
            using (MySqlCommand command = new MySqlCommand())
            {
                connection.Open();
                command.Connection = connection;

                command.CommandText = @"SELECT ВремяВхода, ВремяВыхода 
                                    FROM attendance 
                                    WHERE ФИО = @EmployeeName AND MONTH(ВремяВхода) = @Month AND YEAR(ВремяВхода) = @Year";

                command.Parameters.AddWithValue("@EmployeeName", employeeName);
                command.Parameters.AddWithValue("@Month", month);
                command.Parameters.AddWithValue("@Year", year);

                using (MySqlDataReader reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        DateTime arrivalTime = reader.GetDateTime(0);
                        DateTime departureTime = reader.IsDBNull(1) ? DateTime.Now : reader.GetDateTime(1);

                        TimeSpan timeWorked = departureTime - arrivalTime;
                        totalHoursWorked += timeWorked.TotalHours;
                    }
                }
            }

            return totalHoursWorked;
        }

        private void CountButton_Click(object sender, RoutedEventArgs e)
        {
            int month = Convert.ToInt32(MonthText.Text);
            int year = Convert.ToInt32(YearText.Text);
            string employeeName = FullNameText.Text;

            if (HoursRadioButton.IsChecked == true)
            {
                double totalHoursWorked = CalculateHoursWorkedForMonth(employeeName, month, year);
                MessageBox.Show($"Количество часов работы за месяц: {Convert.ToString(totalHoursWorked)}");
            }
            else if(AbsenseRadioButton.IsChecked == true)
            {
                int daysWithoutAttendance = CalculateDaysWithoutAttendance(employeeName, month, year);
                MessageBox.Show($"Количество дней без посещения: {daysWithoutAttendance}");
            }
        }
        public int CalculateDaysWithoutAttendance(string employeeName, int month, int year)
        {
            int totalDaysInMonth = DateTime.DaysInMonth(year, month);
            int totalAttendanceDays = 0;

            using (MySqlConnection connection = new MySqlConnection(connectionString))
            using (MySqlCommand command = new MySqlCommand())
            {
                connection.Open();
                command.Connection = connection;

                // Подсчет количества записей для сотрудника
                command.CommandText = @"SELECT COUNT(*) 
                                FROM attendance 
                                WHERE ФИО = @EmployeeName AND MONTH(ВремяВхода) = @Month AND YEAR(ВремяВхода) = @Year";

                command.Parameters.AddWithValue("@EmployeeName", employeeName);
                command.Parameters.AddWithValue("@Month", month);
                command.Parameters.AddWithValue("@Year", year);

                totalAttendanceDays = Convert.ToInt32(command.ExecuteScalar());
            }

            // Вычитаем количество дней отсутствия от общего количества дней в месяце
            int daysWithoutAttendance = totalDaysInMonth - totalAttendanceDays-8;

            return daysWithoutAttendance;
        }
        private void ExportToExcelButton_Click(object sender, RoutedEventArgs e)
        {
            string sql = "SELECT * FROM attendance";
            DataTable ettendData = new DataTable();

            using (MySqlConnection connection = new MySqlConnection(connectionString))
            using (MySqlCommand command = new MySqlCommand(sql, connection))
            using (MySqlDataAdapter adapter = new MySqlDataAdapter(command))
            {
                connection.Open();
                adapter.Fill(ettendData);
            }
            foreach (DataRow row in ettendData.Rows)
            {
                // Явно указываем тип данных для столбца ВремяВхода
                ettendData.Columns["ВремяВхода"].DataType = typeof(DateTime);
            }

            // Создание нового Excel-файла
            using (ExcelPackage package = new ExcelPackage())
            {
                // Добавление нового листа в Excel-файл
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Attendance Data");

                // Заполнение листа данными из DataTable
                worksheet.Cells["A1"].LoadFromDataTable(ettendData, true);

                // Сохранение Excel-файла
                string filePath = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "AttendanceData.xlsx");
                FileInfo excelFile = new FileInfo(filePath);
                package.SaveAs(excelFile);
            }

            MessageBox.Show("Данные успешно экспортированы в Excel-файл.");
        }
    }
}


