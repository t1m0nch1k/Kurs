using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace Kurs
{
    public partial class LogViwerWindow : Window
    {
        private List<AppLog> _allLogs;

        public LogViwerWindow()
        {
            InitializeComponent();
            LoadLogs();
        }

        // ---------------------------------------------------------------
        // ЗАГРУЗКА
        // ---------------------------------------------------------------

        private void LoadLogs()
        {
            _allLogs = Logger.ReadAll();

            // Заполняем фильтр-комбо уникальными типами действий
            var actions = _allLogs
                .Select(l => l.Action)
                .Distinct()
                .OrderBy(a => a)
                .ToList();
            actions.Insert(0, "— Все —");
            comboAction.ItemsSource = actions;
            comboAction.SelectedIndex = 0;

            ShowLogs(_allLogs);
            UpdateStatusBar(_allLogs);
        }

        // ---------------------------------------------------------------
        // ОТОБРАЖЕНИЕ
        // ---------------------------------------------------------------

        private void ShowLogs(List<AppLog> logs)
        {
            // Сортируем: новые записи сверху
            gridLogs.ItemsSource = logs.OrderByDescending(l => l.Time).ToList();
            txtCount.Text = $"Показано записей: {logs.Count} из {_allLogs.Count}";
        }

        private void UpdateStatusBar(List<AppLog> logs)
        {
            // Последний вход в систему
            var lastLogin = _allLogs
                .Where(l => l.Action == "Вход в систему")
                .OrderByDescending(l => l.Time)
                .FirstOrDefault();

            txtLastLogin.Text = lastLogin != null
                ? $"Последний вход: {lastLogin.UserName}  в  {lastLogin.Time:dd.MM.yyyy HH:mm}"
                : "Входов не зафиксировано";
        }

        // ---------------------------------------------------------------
        // ФИЛЬТРАЦИЯ
        // ---------------------------------------------------------------

        private void btnFilter_Click(object sender, RoutedEventArgs e)
        {
            var filtered = _allLogs.AsEnumerable();

            // По дате «С»
            if (dpFrom.SelectedDate.HasValue)
                filtered = filtered.Where(l => l.Time >= dpFrom.SelectedDate.Value);

            // По дате «По» (включая конец дня)
            if (dpTo.SelectedDate.HasValue)
                filtered = filtered.Where(l => l.Time < dpTo.SelectedDate.Value.AddDays(1));

            // По имени пользователя
            if (!string.IsNullOrWhiteSpace(txtFilterUser.Text))
            {
                string search = txtFilterUser.Text.Trim().ToLower();
                filtered = filtered.Where(l =>
                    l.UserName != null && l.UserName.ToLower().Contains(search));
            }

            // По типу действия
            if (comboAction.SelectedIndex > 0 && comboAction.SelectedItem is string action)
                filtered = filtered.Where(l => l.Action == action);

            ShowLogs(filtered.ToList());
        }

        private void btnReset_Click(object sender, RoutedEventArgs e)
        {
            dpFrom.SelectedDate = null;
            dpTo.SelectedDate = null;
            txtFilterUser.Text = "";
            comboAction.SelectedIndex = 0;
            ShowLogs(_allLogs);
        }
    }
}