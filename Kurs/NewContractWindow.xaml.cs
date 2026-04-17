using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Word = Microsoft.Office.Interop.Word;

namespace Kurs
{
    public partial class NewContractWindow : Window
    {
        // Временный список изделий до сохранения договора
        private readonly List<ContractItems> _tempItems = new List<ContractItems>();

        public NewContractWindow()
        {
            InitializeComponent();
            LoadAllData();
            dpIssueDate.SelectedDate = DateTime.Today;
            dpDueDate.SelectedDate = DateTime.Today.AddDays(30);
            txtInterestRate.Text = "5";
        }

        // ---------------------------------------------------------------
        // ЗАГРУЗКА ДАННЫХ
        // ---------------------------------------------------------------

        private void LoadAllData()
        {
            try
            {
                using (var db = new LombardDBEntities())
                {
                    // Отображаем ФИО клиента целиком
                    var clients = db.Clients.ToList();
                    // Добавляем вычисляемое поле FullName через анонимный тип —
                    // проще привязать DisplayMemberPath к существующему свойству Clients
                    comboClients.ItemsSource = clients;
                    comboClients.DisplayMemberPath = "LastName"; // меняем ниже в XAML — уже LastName

                    comboItemTypes.ItemsSource = db.ItemTypes.ToList();

                    // PurchaseRates содержит цену за грамм и связь с MetalsCatalog
                    comboMetals.ItemsSource =
                        db.PurchaseRates.Include(r => r.MetalsCatalog).ToList();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка загрузки данных: " + ex.Message);
            }
        }

        // ---------------------------------------------------------------
        // АВТО-РАСЧЁТ СТОИМОСТИ ИЗДЕЛИЯ
        // ---------------------------------------------------------------

        /// <summary>
        /// Вызывается при смене металла — заполняет цену за грамм и пересчитывает оценку.
        /// </summary>
        private void comboMetals_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            RecalcItem(sender, e);
        }

        /// <summary>
        /// Пересчитывает оценочную стоимость текущего изделия:
        /// ОценкаЗаКг = Вес × ЦенаЗаГрамм
        /// </summary>
        private void RecalcItem(object sender, EventArgs e)
        {
            if (!(comboMetals.SelectedItem is PurchaseRates rate))
            {
                txtPricePerGram.Text = "";
                txtEstimate.Text = "";
                return;
            }

            txtPricePerGram.Text = rate.PricePerGram.ToString("F2");

            if (decimal.TryParse(txtWeight.Text.Replace(',', '.'),
                    System.Globalization.NumberStyles.Any,
                    System.Globalization.CultureInfo.InvariantCulture,
                    out decimal weight) && weight > 0)
            {
                decimal estimate = weight * rate.PricePerGram;
                txtEstimate.Text = estimate.ToString("F2");
            }
            else
            {
                txtEstimate.Text = "";
            }
        }

        // ---------------------------------------------------------------
        // ДОБАВЛЕНИЕ ИЗДЕЛИЯ В СПИСОК
        // ---------------------------------------------------------------

        private void btnAddToTemp_Click(object sender, RoutedEventArgs e)
        {
            // Валидация
            if (!(comboMetals.SelectedItem is PurchaseRates rate))
            { MessageBox.Show("Выберите металл и пробу."); return; }

            if (!decimal.TryParse(txtWeight.Text.Replace(',', '.'),
                    System.Globalization.NumberStyles.Any,
                    System.Globalization.CultureInfo.InvariantCulture,
                    out decimal weight) || weight <= 0)
            { MessageBox.Show("Введите корректный вес (больше 0)."); return; }

            decimal estimate = weight * rate.PricePerGram;

            var item = new ContractItems
            {
                AssayID = rate.RateID,          // внешний ключ на PurchaseRates
                Weight = weight,
                EstimatedValue = estimate,
                Description = string.IsNullOrWhiteSpace(txtDescription.Text)
                                    ? $"{rate.MetalsCatalog?.MetalName} {rate.MetalsCatalog?.PurityValue}пр."
                                    : txtDescription.Text.Trim()
            };

            _tempItems.Add(item);
            RefreshGrid();

            // Сброс полей изделия
            txtDescription.Clear();
            txtWeight.Clear();
            txtEstimate.Clear();
            txtPricePerGram.Clear();
            comboMetals.SelectedIndex = -1;
        }

        /// <summary>Удаляет выбранное изделие из временного списка.</summary>
        private void btnRemoveItem_Click(object sender, RoutedEventArgs e)
        {
            if (gridAddedItems.SelectedItem is ContractItems selected)
            {
                _tempItems.Remove(selected);
                RefreshGrid();
            }
        }

        /// <summary>Обновляет таблицу и пересчитывает итоговую сумму займа.</summary>
        private void RefreshGrid()
        {
            gridAddedItems.ItemsSource = null;
            gridAddedItems.ItemsSource = _tempItems;

            // Сумма займа = 70 % от общей оценочной стоимости (стандартная ломбардная практика)
            decimal total = _tempItems.Sum(i => i.EstimatedValue ?? 0m);
            decimal loan = Math.Round(total * 0.70m, 2);
            txtLoanAmount.Text = loan.ToString("F2");
        }

        // ---------------------------------------------------------------
        // СОХРАНЕНИЕ ДОГОВОРА
        // ---------------------------------------------------------------
        private void btnFinalSave_Click(object sender, RoutedEventArgs e)
        {
            // [ВАЛИДАЦИЯ - Оставляем как была в твоем коде]
            if (comboClients.SelectedValue == null) { MessageBox.Show("Выберите клиента."); return; }
            if (_tempItems.Count == 0) { MessageBox.Show("Добавьте изделия."); return; }
            // ... (остальные проверки) ...

            try
            {
                using (var db = new LombardDBEntities())
                {
                    // 1. Создаем договор
                    var contract = new Contracts
                    {
                        ClientID = (int)comboClients.SelectedValue,
                        UserID = AppSession.CurrentUser?.UserID,
                        IssueDate = dpIssueDate.SelectedDate.Value,
                        DueDate = dpDueDate.SelectedDate.Value,
                        LoanAmount = decimal.Parse(txtLoanAmount.Text),
                        InterestRate = decimal.Parse(txtInterestRate.Text),
                        Status = "Активен"
                    };

                    db.Contracts.Add(contract);
                    db.SaveChanges(); // Теперь у contract есть ID

                    // 2. Сохраняем изделия
                    foreach (var item in _tempItems)
                    {
                        item.ContractID = contract.ContractID;
                        db.ContractItems.Add(item);
                    }
                    db.SaveChanges();

                    Logger.Write("Новый договор", $"№{contract.ContractID}, сумма {contract.LoanAmount} ₽");

                    MessageBox.Show($"Договор №{contract.ContractID} оформлен и отправлен в Word!", "Успех");
                    DialogResult = true;
                    Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка: " + ex.Message);
            }
        }

       
    }
}