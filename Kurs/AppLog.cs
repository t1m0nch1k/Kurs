using System;

namespace Kurs
{
    // Класс для хранения записей журнала действий (логов)
    // Логи сохраняются в файл, так как таблицы Logs нет в БД
    public class AppLog
    {
        public DateTime Time { get; set; }       // Время действия
        public string UserName { get; set; }      // Кто выполнил
        public string Action { get; set; }        // Что сделал
        public string Details { get; set; }       // Подробности
    }

    // Статический класс для записи и чтения логов
    public static class Logger
    {
        // Путь к файлу лога — папка рядом с программой
        private static readonly string LogFilePath =
            System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "lombard_log.txt");

        // Записать действие в файл лога
        public static void Write(string action, string details = "")
        {
            try
            {
                // Получаем имя текущего пользователя из сессии
                string userName = AppSession.CurrentUser != null
                    ? AppSession.CurrentUser.FullName
                    : "Неизвестно";

                // Формируем строку лога
                string line = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} | {userName} | {action} | {details}";

                // Добавляем строку в файл (создаём файл если нет)
                System.IO.File.AppendAllText(LogFilePath, line + Environment.NewLine, System.Text.Encoding.UTF8);
            }
            catch
            {
                // Если запись лога не удалась — просто игнорируем,
                // чтобы ошибка логирования не мешала работе программы
            }
        }

        // Прочитать все логи из файла
        public static System.Collections.Generic.List<AppLog> ReadAll()
        {
            var result = new System.Collections.Generic.List<AppLog>();

            if (!System.IO.File.Exists(LogFilePath))
                return result;

            try
            {
                string[] lines = System.IO.File.ReadAllLines(LogFilePath, System.Text.Encoding.UTF8);
                foreach (string line in lines)
                {
                    // Каждая строка формата: "дата | пользователь | действие | подробности"
                    string[] parts = line.Split('|');
                    if (parts.Length >= 3)
                    {
                        result.Add(new AppLog
                        {
                            Time = DateTime.TryParse(parts[0].Trim(), out DateTime dt) ? dt : DateTime.MinValue,
                            UserName = parts.Length > 1 ? parts[1].Trim() : "",
                            Action = parts.Length > 2 ? parts[2].Trim() : "",
                            Details = parts.Length > 3 ? parts[3].Trim() : ""
                        });
                    }
                }
            }
            catch { }

            // Возвращаем от новых к старым
            result.Reverse();
            return result;
        }
    }
}
