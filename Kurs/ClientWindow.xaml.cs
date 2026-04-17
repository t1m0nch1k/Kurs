using System;
using System.Data.SqlClient;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;

namespace Kurs
{
    public partial class ClientWindow : Window
    {
        public Clients Client { get; private set; }
        private readonly bool _isNew;

        public ClientWindow(Clients client = null)
        {
            InitializeComponent();
            _isNew = (client == null);
            Client = client ?? new Clients();

            txtTitle.Text = _isNew ? "Новый клиент" : "Редактирование клиента";

            // Заполняем поля вручную (без DataContext — избегаем конфликтов валидации)
            txtLastName.Text = Client.LastName ?? "";
            txtFirstName.Text = Client.FirstName ?? "";
            txtMiddleName.Text = Client.MiddleName ?? "";
            txtPhone.Text = Client.Phone ?? "";
            txtPassport.Text = Client.Passport ?? "";
            dpBirthDate.SelectedDate = Client.BirthDate;
        }

        // ---------------------------------------------------------------
        // АВТО-ВОЗРАСТ
        // ---------------------------------------------------------------

        private void dpBirthDate_Changed(object sender, SelectionChangedEventArgs e)
        {
            if (dpBirthDate.SelectedDate.HasValue)
            {
                int age = CalcAge(dpBirthDate.SelectedDate.Value);
                txtAge.Text = $"{age} лет";
                txtAge.Foreground = age < 18
                    ? System.Windows.Media.Brushes.Red
                    : System.Windows.Media.Brushes.DarkGreen;
            }
            else
            {
                txtAge.Text = "";
            }
        }

        private static int CalcAge(DateTime birth)
        {
            int age = DateTime.Today.Year - birth.Year;
            if (birth.Date > DateTime.Today.AddYears(-age)) age--;
            return age;
        }

        // ---------------------------------------------------------------
        // СОХРАНЕНИЕ
        // ---------------------------------------------------------------

        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            HideError();

            // --- обязательные поля ---
            if (string.IsNullOrWhiteSpace(txtLastName.Text))
            { ShowError("Введите фамилию клиента."); return; }

            if (string.IsNullOrWhiteSpace(txtFirstName.Text))
            { ShowError("Введите имя клиента."); return; }

            if (string.IsNullOrWhiteSpace(txtPhone.Text))
            { ShowError("Введите номер телефона."); return; }

            if (string.IsNullOrWhiteSpace(txtPassport.Text))
            { ShowError("Введите паспортные данные."); return; }

            if (dpBirthDate.SelectedDate == null)
            { ShowError("Укажите дату рождения."); return; }

            // --- возраст 18+ ---
            int age = CalcAge(dpBirthDate.SelectedDate.Value);
            if (age < 18)
            {
                ShowError($"Клиенту {age} лет — регистрация разрешена только с 18 лет.");
                return;
            }

            // --- телефон: цифры/пробелы/+/- ---
            if (!Regex.IsMatch(txtPhone.Text.Trim(), @"^[\d\s\+\-\(\)]{7,20}$"))
            { ShowError("Телефон содержит недопустимые символы или слишком короткий."); return; }

            // --- запись в объект ---
            Client.LastName = txtLastName.Text.Trim();
            Client.FirstName = txtFirstName.Text.Trim();
            Client.MiddleName = txtMiddleName.Text.Trim();
            Client.Phone = txtPhone.Text.Trim();
            Client.Passport = txtPassport.Text.Trim();
            Client.BirthDate = dpBirthDate.SelectedDate.Value;

            Logger.Write(
                _isNew ? "Добавлен клиент" : "Редактирование клиента",
                $"{Client.LastName} {Client.FirstName}");

            DialogResult = true;
        }


        private void ShowError(string msg)
        {
            txtError.Text = msg;
            pnlError.Visibility = Visibility.Visible;
        }

        private void HideError() =>
            pnlError.Visibility = Visibility.Collapsed;
    }
}