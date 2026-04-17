using System.Data.SqlClient;
using System.Linq;
using System.Windows;

namespace Kurs
{
    public partial class UserEditWindow : Window
    {
        public Users Employee { get; private set; }
        private readonly bool _isNew;

        public UserEditWindow(Users user = null)
        {
            InitializeComponent();
            _isNew = (user == null);
            Employee = user ?? new Users();
            DataContext = Employee;

            try
            {
                using (var db = new LombardDBEntities())
                {
                    cbRole.ItemsSource = db.UserRoles.ToList();
                    if (Employee.RoleID.HasValue)
                        cbRole.SelectedValue = Employee.RoleID;
                }
            }
            catch { }
        }

        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            pnlError.Visibility = Visibility.Collapsed;

            if (string.IsNullOrWhiteSpace(txtFullName.Text))
            { ShowError("Введите полное имя."); return; }

            if (string.IsNullOrWhiteSpace(txtLogin.Text))
            { ShowError("Введите логин."); return; }

            if (_isNew && string.IsNullOrWhiteSpace(pbPassword.Password))
            { ShowError("Введите пароль."); return; }

            Employee.FullName = txtFullName.Text.Trim();
            Employee.Login = txtLogin.Text.Trim();

            if (!string.IsNullOrWhiteSpace(pbPassword.Password))
                Employee.Password = pbPassword.Password;

            if (cbRole.SelectedItem is UserRoles role)
                Employee.RoleID = role.RoleID;

            Logger.Write(
                _isNew ? "Добавление пользователя" : "Редактирование пользователя",
                Employee.FullName);

            DialogResult = true;
        }

        private void ShowError(string msg)
        {
            txtError.Text = msg;
            pnlError.Visibility = Visibility.Visible;
        }
    }
}