using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
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
    public partial class LoginWindow : Window
    {
        private string connectionString;
        public LoginWindow()
        {
            InitializeComponent();
            connectionString = "Server=localhost;Database=empattend;Uid=root;Pwd=Kate+Kate19;";
        }
        private void LogButton_Click(object sender, RoutedEventArgs e)
        {
            string username = LoginBox.Text;
            string password = PassBox.Password;

            LogUser(username);
            // Проверка логина и пароля
            if (CheckCredentials(username, password))
            {
                // Открытие главного окна
                MainWindow mainWindow = new MainWindow();
                mainWindow.Show();
                this.Close();
            }
            else
            {
                // Если логин и пароль неверны, показывается сообщение об ошибке
                MessageBox.Show("Пароль неверный");
            }
            
        }

        private bool CheckCredentials(string username, string password)
        {
            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                connection.Open();

                string query = "SELECT COUNT(*) FROM logins WHERE username = @username AND password = @password";
                MySqlCommand command = new MySqlCommand(query, connection);
                command.Parameters.AddWithValue("@username", username);
                command.Parameters.AddWithValue("@password", password);

                // Выполнение запроса и получение количества совпадений
                int count = Convert.ToInt32(command.ExecuteScalar());

                // Если количество совпадений больше 0, то логин и пароль верны
                return count > 0;
            }
        }
        private void LogUser(string username)
        {
            string logFilePath = "C:\\Users\\katek\\source\\c#\\AttendanceEmp\\Log.txt";
            //Создание записи о времени входа пользователя в текстовом файле
            string logMessage = $"{username} вошёл в систему {DateTime.Now}";
            //Внесение записи в текстовый файл
            using (StreamWriter writer = new StreamWriter(logFilePath, true))
            {
                writer.WriteLine(logMessage);
            }
        }
    }
}
