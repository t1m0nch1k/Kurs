using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Data.Entity;

namespace Kurs
{
    // Класс для генерации договора в формате Word (.docx) и PDF
    // Используем встроенные возможности Windows: XPS и Word XML
    public static class ContractPrinter
    {
        // Сгенерировать договор в формате HTML (универсально открывается в Word и браузере)
        // Word умеет открывать HTML-файлы с расширением .doc
        public static void PrintToWord(Contracts contract, List<ContractItems> items)
        {
            try
            {
                // Запрашиваем путь для сохранения
                var dialog = new Microsoft.Win32.SaveFileDialog
                {
                    Title = "Сохранить договор",
                    Filter = "Word документ (*.doc)|*.doc",
                    FileName = $"Договор_№{contract.ContractID}_{DateTime.Now:dd-MM-yyyy}"
                };

                if (dialog.ShowDialog() != true) return;

                string html = BuildHtml(contract, items);
                File.WriteAllText(dialog.FileName, html, Encoding.UTF8);

                MessageBox.Show($"Договор сохранён:\n{dialog.FileName}\n\nОткройте файл в Microsoft Word.",
                    "Готово", MessageBoxButton.OK, MessageBoxImage.Information);

                Logger.Write("Печать договора", $"Договор №{contract.ContractID} сохранён в Word");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при сохранении в Word: " + ex.Message);
            }
        }

        // Сгенерировать договор в формате PDF через XPS
        // XPS — встроенный в Windows формат; мы создаём HTML и просим пользователя использовать «Печать в PDF»
        public static void PrintToPdf(Contracts contract, List<ContractItems> items)
        {
            try
            {
                // Сохраняем HTML-версию договора, затем открываем в браузере
                // Пользователь нажимает Ctrl+P и выбирает «Сохранить как PDF»
                string tempPath = Path.Combine(Path.GetTempPath(),
                    $"Договор_{contract.ContractID}_{DateTime.Now:HHmmss}.html");

                string html = BuildHtml(contract, items);
                File.WriteAllText(tempPath, html, Encoding.UTF8);

                // Открываем в браузере по умолчанию
                System.Diagnostics.Process.Start(tempPath);

                MessageBox.Show(
                    "Договор открыт в браузере.\n\n" +
                    "Для сохранения в PDF:\n" +
                    "1. Нажмите Ctrl+P (Печать)\n" +
                    "2. Выберите принтер «Microsoft Print to PDF»\n" +
                    "3. Нажмите «Печать» и укажите путь для сохранения",
                    "Сохранение в PDF",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);

                Logger.Write("Печать договора", $"Договор №{contract.ContractID} открыт для печати в PDF");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при открытии для печати: " + ex.Message);
            }
        }

        // Формирует HTML-текст договора
        private static string BuildHtml(Contracts contract, List<ContractItems> items)
        {
            // Собираем имя клиента
            string clientName = contract.Clients != null
                ? $"{contract.Clients.LastName} {contract.Clients.FirstName} {contract.Clients.MiddleName}"
                : "—";

            string passport = contract.Clients?.Passport ?? "—";
            string phone = contract.Clients?.Phone ?? "—";

            // Собираем имя сотрудника
            string userName = contract.Users?.FullName ?? "—";

            // Формируем строки таблицы изделий
            var rowsBuilder = new StringBuilder();
            int num = 1;
            decimal totalEstimate = 0;

            foreach (var ci in items)
            {
                totalEstimate += ci.EstimatedValue ?? 0;
                rowsBuilder.AppendLine($@"
                <tr>
                    <td style='padding:6px;border:1px solid #ccc;text-align:center;'>{num++}</td>
                    <td style='padding:6px;border:1px solid #ccc;'>{ci.Description ?? "—"}</td>
                    <td style='padding:6px;border:1px solid #ccc;text-align:center;'>{ci.Weight} гр</td>
                    <td style='padding:6px;border:1px solid #ccc;text-align:right;'>{ci.EstimatedValue:N0} ₽</td>
                </tr>");
            }

            // Процентная ставка и срок
            string interestRate = contract.InterestRate.ToString("0.##") + "% в месяц";
            string issueDate = contract.IssueDate?.ToString("dd.MM.yyyy") ?? "—";
            string dueDate = contract.DueDate.ToString("dd.MM.yyyy");

            return $@"<!DOCTYPE html>
<html lang='ru'>
<head>
<meta charset='UTF-8'/>
<title>Договор №{contract.ContractID}</title>
<style>
  body      {{ font-family: 'Times New Roman', serif; font-size: 14px; margin: 40px; color: #000; }}
  h2        {{ text-align: center; font-size: 18px; margin-bottom: 5px; }}
  .subtitle {{ text-align: center; font-size: 13px; margin-bottom: 30px; }}
  .section  {{ margin-bottom: 15px; }}
  table     {{ width: 100%; border-collapse: collapse; margin: 10px 0; }}
  th        {{ background: #f0f0f0; padding: 6px; border: 1px solid #ccc; font-size: 13px; }}
  td        {{ font-size: 13px; }}
  .sign-row {{ display: flex; justify-content: space-between; margin-top: 40px; }}
  .sign-box {{ width: 45%; }}
  .underline{{ border-bottom: 1px solid #000; height: 20px; margin-top: 5px; }}
  .bold     {{ font-weight: bold; }}
  .right    {{ text-align: right; }}
  @media print {{
    body {{ margin: 20px; }}
  }}
</style>
</head>
<body>

<h2>ЛОМБАРДНАЯ СЕТЬ «АЛТЫН»</h2>
<p class='subtitle'>ДОГОВОР ЗАЛОГА №{contract.ContractID}</p>

<div class='section'>
  <p>г. ________________ &nbsp;&nbsp;&nbsp; <span class='bold'>Дата выдачи: {issueDate}</span>
     &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <span class='bold'>Срок возврата: {dueDate}</span></p>
</div>

<div class='section'>
  <p class='bold'>1. СТОРОНЫ ДОГОВОРА</p>
  <p><b>Ломбард:</b> ООО «АЛТЫН», сотрудник: {userName}</p>
  <p><b>Клиент:</b> {clientName}</p>
  <p><b>Паспорт:</b> {passport} &nbsp;&nbsp;&nbsp; <b>Телефон:</b> {phone}</p>
</div>

<div class='section'>
  <p class='bold'>2. ПРЕДМЕТ ЗАЛОГА</p>
  <table>
    <thead>
      <tr>
        <th style='width:40px;'>№</th>
        <th>Описание изделия</th>
        <th style='width:90px;'>Вес</th>
        <th style='width:120px;'>Оценка</th>
      </tr>
    </thead>
    <tbody>
      {rowsBuilder}
    </tbody>
    <tfoot>
      <tr>
        <td colspan='3' style='padding:6px;border:1px solid #ccc;text-align:right;font-weight:bold;'>Итого оценка:</td>
        <td style='padding:6px;border:1px solid #ccc;text-align:right;font-weight:bold;'>{totalEstimate:N0} ₽</td>
      </tr>
    </tfoot>
  </table>
</div>

<div class='section'>
  <p class='bold'>3. УСЛОВИЯ ЗАЙМА</p>
  <p><b>Сумма займа:</b> {contract.LoanAmount:N0} ₽</p>
  <p><b>Процентная ставка:</b> {interestRate}</p>
  <p><b>Срок погашения:</b> {dueDate}</p>
  <p>Клиент обязуется погасить задолженность до указанной даты.
     В случае непогашения в срок ломбард вправе реализовать предмет залога.</p>
</div>

<div class='section'>
  <p class='bold'>4. ПОДПИСИ СТОРОН</p>
  <div class='sign-row'>
    <div class='sign-box'>
      <p>Сотрудник ломбарда:</p>
      <p>{userName}</p>
      <div class='underline'></div>
      <p style='font-size:12px;color:#555;'>подпись / дата</p>
    </div>
    <div class='sign-box' style='text-align:right;'>
      <p>Клиент:</p>
      <p>{clientName}</p>
      <div class='underline'></div>
      <p style='font-size:12px;color:#555;'>подпись / дата</p>
    </div>
  </div>
</div>

</body>
</html>";
        }
    }
}
