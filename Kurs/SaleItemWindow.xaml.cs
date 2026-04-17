using System;
using System.Data.Entity;
using System.Linq;
using System.Windows;

namespace Kurs
{
    public partial class SaleItemWindow : Window
    {
        // Договор, из которого продаём изделия
        private Contracts _contract;

        public SaleItemWindow(Contracts contract)
        {
            InitializeComponent();
            _contract = contract;
            LoadData();
        }

        // Загружаем информацию о договоре и список его изделий
        private void LoadData()
        {
            // Показываем краткую информацию о договоре
            txtContractInfo.Text = $"Договор №{_contract.ContractID} | " +
                                   $"Клиент: {_contract.Clients?.LastName} | " +
                                   $"Просрочен до: {_contract.DueDate:dd.MM.yyyy}";

            // Загружаем изделия этого договора
            try
            {
                using (var db = new LombardDBEntities())
                {
                    // Получаем ContractItems для этого договора и связанные Items
                    var contractItems = db.ContractItems
                        .Include(ci => ci.ItemID)
                        .Where(ci => ci.ContractID == _contract.ContractID)
                        .ToList();

                    // Показываем в выпадающем списке Items (изделия)
                    comboItems.ItemsSource = contractItems;
                    comboItems.DisplayMemberPath = null; // используем ItemTemplate

                    if (contractItems.Count == 0)
                    {
                        MessageBox.Show("В этом договоре нет изделий для продажи.");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка загрузки изделий: " + ex.Message);
            }
        }

        // Кнопка «Оформить продажу»
        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            // Проверяем, что изделие выбрано
            if (comboItems.SelectedItem == null)
            {
                MessageBox.Show("Выберите изделие из списка!");
                return;
            }

            // Проверяем имя покупателя
            if (string.IsNullOrWhiteSpace(txtBuyerName.Text))
            {
                MessageBox.Show("Введите ФИО покупателя!");
                return;
            }

            // Проверяем, что цена — корректное число
            if (!decimal.TryParse(txtSalePrice.Text, out decimal salePrice) || salePrice <= 0)
            {
                MessageBox.Show("Введите корректную цену продажи (число больше нуля)!");
                return;
            }

            try
            {
                using (var db = new LombardDBEntities())
                {
                    // Получаем выбранное изделие из договора
                    var selectedContractItem = (ContractItems)comboItems.SelectedItem;

                    // Создаём запись о продаже в таблице Sales
                    var sale = new Sales
                    {
                        ContractID = _contract.ContractID,
                        ItemID = selectedContractItem.ItemID,
                        SaleDate = DateTime.Now,
                        SalePrice = salePrice,
                        BuyerName = txtBuyerName.Text.Trim()
                    };

                    db.Sales.Add(sale);

                    // Меняем статус договора на «Продан»
                    var contract = db.Contracts.Find(_contract.ContractID);
                    if (contract != null)
                    {
                        contract.Status = "Продан";
                        db.Entry(contract).State = EntityState.Modified;
                    }

                    db.SaveChanges();

                    // Записываем в лог
                    Logger.Write("Продажа изделия",
                        $"Договор №{_contract.ContractID}, покупатель: {txtBuyerName.Text}, сумма: {salePrice} ₽");

                    MessageBox.Show($"Продажа успешно оформлена!\nПокупатель: {txtBuyerName.Text}\nСумма: {salePrice:N0} ₽",
                        "Успех", MessageBoxButton.OK, MessageBoxImage.Information);

                    this.DialogResult = true;
                    this.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при оформлении продажи: " + ex.Message);
            }
        }
    }
}
