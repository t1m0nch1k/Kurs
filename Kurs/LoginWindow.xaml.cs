using System;
using System.Linq;
using System.Windows;

namespace Kurs
{
    public partial class LoginWindow : Window
    {
        public LoginWindow()
        {
            InitializeComponent();
        }

        private void Login_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                using (var db = new LombardDBEntities())
                {
                    // Ищем пользователя по логину и паролю
                    var user = db.Users.FirstOrDefault(u => u.Login == txtLogin.Text && u.Password == txtPassword.Password);

                    if (user != null)
                    {
                        // Сохраняем данные пользователя в сессию и открываем главное окно
                        AppSession.CurrentUser = user;
                        MainWindow main = new MainWindow();
                        main.Show();
                        this.Close();
                    }
                    else
                    {
                        txtError.Text = "Неверный логин или пароль!";
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка подключения к БД: " + ex.Message);
            }
        }
    }
}