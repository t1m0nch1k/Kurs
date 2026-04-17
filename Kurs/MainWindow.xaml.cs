using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Data.Entity;
using System.Collections.Generic;
using System.IO;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;

namespace Kurs
{
    public partial class MainWindow : Window
    {
        // ===== ПОЛЯ ДЛЯ ПАГИНАЦИИ СКЛАДА =====
        private int _warehouseCurrentPage = 1;
        private int _warehousePageSize = 10;
        private int _warehouseTotalItems = 0;
        private int _warehouseTotalPages = 1;
        private List<Items> _allWarehouseItems = new List<Items>();

        // ===== ПОЛЯ ДЛЯ ФИЛЬТРОВ =====
        private string _contractSearch = "";
        private string _contractStatusFilter = "";
        private string _clientSearch = "";
        private string _clientPhoneSearch = "";
        private string _clientSort = "LastName ASC";
        private string _itemSearch = "";
        private int? _itemTypeFilter = null;
        private int? _itemMetalFilter = null;
        private string _userSearch = "";
        private int? _userRoleFilter = null;

        public MainWindow()
        {
            InitializeComponent();
            ApplyRolePermissions();
            LoadDataFromDb();

            this.Loaded += (s, e) => { if (btnHome != null) btnHome.Focus(); };
        }

        private void ApplyRolePermissions()
        {
            var user = AppSession.CurrentUser;
            if (user == null) return;

            if (user.UserRoles != null)
                txtUserInfo.Text = $"👤 {user.FullName}\nДоступ: {user.UserRoles.RoleName}";

            if (user.RoleID != 3 && btnAdmin != null)
                btnAdmin.Visibility = Visibility.Collapsed;
        }

        private void LoadDataFromDb()
        {
            try
            {
                using (var db = new LombardDBEntities())
                {
                    var allContracts = db.Contracts.Include(c => c.Clients).ToList();
                    if (gridContracts != null)
                        gridContracts.ItemsSource = allContracts;

                    var activeContracts = allContracts.Where(c => c.Status == "Активен").ToList();

                    if (txtActiveContracts != null)
                        txtActiveContracts.Text = activeContracts.Count.ToString();

                    if (txtTotalSum != null)
                    {
                        decimal totalSum = activeContracts.Sum(c => (decimal?)c.LoanAmount) ?? 0;
                        txtTotalSum.Text = totalSum.ToString("N0") + " ₽";
                    }

                    if (gridPurchaseRates != null)
                        gridPurchaseRates.ItemsSource =
                            db.PurchaseRates.Include(r => r.MetalsCatalog).ToList();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка загрузки данных: " + ex.Message);
            }
        }

        private void OpenHome(object sender, RoutedEventArgs e)
        {
            MainContent.SelectedIndex = 0;
            LoadDataFromDb();
        }

        private void OpenContracts(object sender, RoutedEventArgs e)
        {
            MainContent.SelectedIndex = 1;
            LoadContractsWithFilters();
        }

        private void OpenClients(object sender, RoutedEventArgs e)
        {
            MainContent.SelectedIndex = 2;
            LoadClientsWithFilters();
        }

        private void OpenWarehouse(object sender, RoutedEventArgs e)
        {
            MainContent.SelectedIndex = 3;
            LoadItemFilters();
            LoadWarehouseItems();
        }

        private void OpenAdmin(object sender, RoutedEventArgs e)
        {
            MainContent.SelectedIndex = 4;
            LoadUserRoles();
            LoadUsersWithFilters();
        }

        // ═══════════════════════════════════════════════════════════════
        //                    КОНТРАКТЫ - ПОИСК И ФИЛЬТРЫ
        // ═══════════════════════════════════════════════════════════════

        private void txtSearchContracts_TextChanged(object sender, TextChangedEventArgs e)
        {
            _contractSearch = txtSearchContracts.Text.ToLower();
            LoadContractsWithFilters();
        }

        private void cmbContractStatus_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (cmbContractStatus.SelectedItem is ComboBoxItem item)
            {
                _contractStatusFilter = item.Content.ToString() == "Все статусы" ? "" : item.Content.ToString();
                LoadContractsWithFilters();
            }
        }

        private void btnResetContractFilters_Click(object sender, RoutedEventArgs e)
        {
            txtSearchContracts.Text = "";
            cmbContractStatus.SelectedIndex = 0;
            LoadContractsWithFilters();
        }

        private void LoadContractsWithFilters()
        {
            try
            {
                using (var db = new LombardDBEntities())
                {
                    var query = db.Contracts.Include(c => c.Clients).AsQueryable();

                    if (!string.IsNullOrEmpty(_contractSearch))
                    {
                        query = query.Where(c => c.Clients.LastName.ToLower().Contains(_contractSearch) ||
                                                c.Clients.FirstName.ToLower().Contains(_contractSearch));
                    }

                    if (!string.IsNullOrEmpty(_contractStatusFilter))
                    {
                        query = query.Where(c => c.Status == _contractStatusFilter);
                    }

                    // Добавлена проверка на null
                    if (gridContracts != null)
                    {
                        gridContracts.ItemsSource = query.ToList();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка загрузки контрактов: " + ex.Message);
            }
        }

        // ═══════════════════════════════════════════════════════════════
        //                    КЛИЕНТЫ - ПОИСК И ФИЛЬТРЫ
        // ═══════════════════════════════════════════════════════════════

        private void txtSearchClients_TextChanged(object sender, TextChangedEventArgs e)
        {
            _clientSearch = txtSearchClients.Text.ToLower();
            LoadClientsWithFilters();
        }

        private void txtSearchClientPhone_TextChanged(object sender, TextChangedEventArgs e)
        {
            _clientPhoneSearch = txtSearchClientPhone.Text;
            LoadClientsWithFilters();
        }

        private void cmbClientSort_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (cmbClientSort.SelectedItem is ComboBoxItem item)
            {
                _clientSort = item.Content.ToString();
                LoadClientsWithFilters();
            }
        }

        private void btnResetClientFilters_Click(object sender, RoutedEventArgs e)
        {
            txtSearchClients.Text = "";
            txtSearchClientPhone.Text = "";
            cmbClientSort.SelectedIndex = 0;
            LoadClientsWithFilters();
        }

        private void LoadClientsWithFilters()
        {
            try
            {
                using (var db = new LombardDBEntities())
                {
                    var query = db.Clients.AsQueryable();

                    if (!string.IsNullOrEmpty(_clientSearch))
                    {
                        query = query.Where(c => c.LastName.ToLower().Contains(_clientSearch) ||
                                                c.FirstName.ToLower().Contains(_clientSearch));
                    }

                    if (!string.IsNullOrEmpty(_clientPhoneSearch))
                    {
                        query = query.Where(c => c.Phone.Contains(_clientPhoneSearch));
                    }

                    switch (_clientSort)
                    {
                        case "По фамилии (А-Я)":
                            query = query.OrderBy(c => c.LastName);
                            break;
                        case "По фамилии (Я-А)":
                            query = query.OrderByDescending(c => c.LastName);
                            break;
                        case "По дате добавления":
                            query = query.OrderByDescending(c => c.ClientID);
                            break;
                        default:
                            query = query.OrderBy(c => c.LastName);
                            break;
                    }

                    // Добавлена проверка на null
                    if (gridClients != null)
                    {
                        gridClients.ItemsSource = query.ToList();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка загрузки клиентов: " + ex.Message);
            }
        }

        private void RefreshClients_Click(object sender, RoutedEventArgs e) =>
            LoadClientsWithFilters();

        private void btnAddClient_Click(object sender, RoutedEventArgs e)
        {
            var win = new ClientWindow();
            if (win.ShowDialog() == true)
            {
                using (var db = new LombardDBEntities())
                {
                    db.Clients.Add(win.Client);
                    try
                    {
                        db.SaveChanges();
                    }
                    catch (System.Data.Entity.Validation.DbEntityValidationException dbEx)
                    {
                        foreach (var ve in dbEx.EntityValidationErrors)
                            foreach (var err in ve.ValidationErrors)
                                MessageBox.Show($"Свойство: {err.PropertyName}  Ошибка: {err.ErrorMessage}");
                    }
                }
                LoadClientsWithFilters();
                Logger.Write("Добавлен клиент", $"{win.Client.LastName} {win.Client.FirstName}");
            }
        }

        private void btnEditClient_Click(object sender, RoutedEventArgs e)
        {
            if (gridClients.SelectedItem is Clients selected)
            {
                var win = new ClientWindow(selected);
                if (win.ShowDialog() == true)
                {
                    using (var db = new LombardDBEntities())
                    {
                        db.Entry(selected).State = EntityState.Modified;
                        db.SaveChanges();
                    }
                    LoadClientsWithFilters();
                }
            }
            else
            {
                MessageBox.Show("Выберите клиента для редактирования.");
            }
        }

        // ═══════════════════════════════════════════════════════════════
        //                    СКЛАД - ПОИСК, ФИЛЬТРЫ И ПАГИНАЦИЯ
        // ═══════════════════════════════════════════════════════════════

        private void LoadWarehouseItems(int pageNumber = 1, int pageSize = 10)
        {
            try
            {
                using (var db = new LombardDBEntities())
                {
                    var query = db.Items.Include(x => x.MetalsCatalog).Include(x => x.ItemTypes).AsQueryable();

                    if (!string.IsNullOrEmpty(_itemSearch))
                    {
                        query = query.Where(i => i.ItemDescription.ToLower().Contains(_itemSearch.ToLower()));
                    }

                    if (_itemTypeFilter.HasValue)
                    {
                        query = query.Where(i => i.TypeID == _itemTypeFilter);
                    }

                    if (_itemMetalFilter.HasValue)
                    {
                        query = query.Where(i => i.AssayID == _itemMetalFilter);
                    }

                    _warehouseTotalItems = query.Count();
                    _warehouseTotalPages = _warehouseTotalItems == 0 ? 1 : (int)Math.Ceiling((double)_warehouseTotalItems / pageSize);

                    if (pageNumber < 1) pageNumber = 1;
                    if (pageNumber > _warehouseTotalPages && _warehouseTotalPages > 0)
                        pageNumber = _warehouseTotalPages;

                    _warehouseCurrentPage = pageNumber;
                    _warehousePageSize = pageSize;

                    var items = query
                        .OrderBy(x => x.ItemID)
                        .Skip((pageNumber - 1) * pageSize)
                        .Take(pageSize)
                        .ToList();

                    // Отображение изделий в виде плиток как на маркетплейсе
                    if (panelItems != null)
                    {
                        panelItems.Children.Clear();
                        foreach (var item in items)
                        {
                            var card = CreateItemCard(item);
                            panelItems.Children.Add(card);
                        }
                    }

                    UpdateWarehousePaginationUI();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка загрузки данных склада: " + ex.Message);
            }
        }

        /// <summary>
        /// Создает карточку изделия для отображения в стиле маркетплейса
        /// </summary>
        private System.Windows.Controls.Border CreateItemCard(Items item)
        {
            var border = new System.Windows.Controls.Border
            {
                Width = 200,
                Height = 280,
                Margin = new System.Windows.Thickness(10),
                Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(255, 255, 255)),
                BorderBrush = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(220, 220, 220)),
                BorderThickness = new System.Windows.Thickness(1),
                CornerRadius = new System.Windows.CornerRadius(8)
            };

            var stackPanel = new System.Windows.Controls.StackPanel
            {
                Margin = new System.Windows.Thickness(10)
            };

            // Изображение изделия
            var image = new System.Windows.Controls.Image
            {
                Width = 180,
                Height = 180,
                Stretch = System.Windows.Media.Stretch.UniformToFill,
                Margin = new System.Windows.Thickness(0, 0, 0, 10)
            };

            if (item.ItemPhoto != null && item.ItemPhoto.Length > 0)
            {
                var bitmap = new System.Windows.Media.Imaging.BitmapImage();
                bitmap.BeginInit();
                bitmap.StreamSource = new System.IO.MemoryStream(item.ItemPhoto);
                bitmap.CacheOption = System.Windows.Media.Imaging.BitmapCacheOption.OnLoad;
                bitmap.EndInit();
                bitmap.Freeze();
                image.Source = bitmap;
            }
            else
            {
                // Заглушка если нет фото
                image.Source = null;
            }

            stackPanel.Children.Add(image);

            // Описание
            var descText = new System.Windows.Controls.TextBlock
            {
                Text = item.ItemDescription ?? "Без описания",
                FontWeight = System.Windows.FontWeights.SemiBold,
                FontSize = 14,
                TextWrapping = System.Windows.TextWrapping.Wrap,
                MaxHeight = 40,
                Margin = new System.Windows.Thickness(0, 0, 0, 5)
            };
            stackPanel.Children.Add(descText);

            // Вес
            var weightText = new System.Windows.Controls.TextBlock
            {
                Text = $"Вес: {item.Weight:F2} г",
                FontSize = 12,
                Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(100, 100, 100)),
                Margin = new System.Windows.Thickness(0, 0, 0, 3)
            };
            stackPanel.Children.Add(weightText);

            // Цена
            var priceText = new System.Windows.Controls.TextBlock
            {
                Text = $"{item.EstimatedValue:F0} ₽",
                FontSize = 16,
                FontWeight = System.Windows.FontWeights.Bold,
                Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(212, 175, 55)), // Золотой цвет
                Margin = new System.Windows.Thickness(0, 3, 0, 0)
            };
            stackPanel.Children.Add(priceText);

            border.Child = stackPanel;
            
            // Добавляем обработчик клика для просмотра деталей
            border.MouseLeftButtonUp += (s, e) =>
            {
                var detailWindow = new ItemDetailWindow(item);
                detailWindow.Owner = this;
                detailWindow.ShowDialog();
            };

            return border;
        }

        private void UpdateWarehousePaginationUI()
        {
            if (txtCurrentPage != null)
                txtCurrentPage.Text = _warehouseCurrentPage.ToString();

            if (txtTotalPages != null)
                txtTotalPages.Text = _warehouseTotalPages.ToString();

            if (btnFirst != null)
                btnFirst.IsEnabled = _warehouseCurrentPage > 1;

            if (btnPrev != null)
                btnPrev.IsEnabled = _warehouseCurrentPage > 1;

            if (btnNext != null)
                btnNext.IsEnabled = _warehouseCurrentPage < _warehouseTotalPages;

            if (btnLast != null)
                btnLast.IsEnabled = _warehouseCurrentPage < _warehouseTotalPages;

            if (txtItemsInfo != null)
            {
                int from = _warehouseTotalItems == 0 ? 0 : (_warehouseCurrentPage - 1) * _warehousePageSize + 1;
                int to = Math.Min(_warehouseCurrentPage * _warehousePageSize, _warehouseTotalItems);
                txtItemsInfo.Text = $"Записи {from}-{to} из {_warehouseTotalItems}";
            }
        }

        private void txtSearchItems_TextChanged(object sender, TextChangedEventArgs e)
        {
            _itemSearch = txtSearchItems.Text;
            LoadWarehouseItems(1, _warehousePageSize);
        }

        private void cmbItemType_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (cmbItemType.SelectedItem is ComboBoxItem item)
            {
                if (item.Content.ToString() == "Все типы")
                    _itemTypeFilter = null;
                else if (item.Tag != null && int.TryParse(item.Tag.ToString(), out int typeId))
                    _itemTypeFilter = typeId;
                LoadWarehouseItems(1, _warehousePageSize);
            }
        }

        private void cmbItemMetal_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (cmbItemMetal.SelectedItem is ComboBoxItem item)
            {
                if (item.Content.ToString() == "Все металлы")
                    _itemMetalFilter = null;
                else if (item.Tag != null && int.TryParse(item.Tag.ToString(), out int metalId))
                    _itemMetalFilter = metalId;
                LoadWarehouseItems(1, _warehousePageSize);
            }
        }

        private void btnResetItemFilters_Click(object sender, RoutedEventArgs e)
        {
            txtSearchItems.Text = "";
            cmbItemType.SelectedIndex = 0;
            cmbItemMetal.SelectedIndex = 0;
            LoadWarehouseItems(1, _warehousePageSize);
        }

        private void LoadItemFilters()
        {
            try
            {
                using (var db = new LombardDBEntities())
                {
                    cmbItemType.Items.Clear();
                    cmbItemType.Items.Add(new ComboBoxItem { Content = "Все типы" });
                    foreach (var type in db.ItemTypes.ToList())
                    {
                        cmbItemType.Items.Add(new ComboBoxItem
                        {
                            Content = type.TypeName,
                            Tag = type.TypeID.ToString()
                        });
                    }

                    cmbItemMetal.Items.Clear();
                    cmbItemMetal.Items.Add(new ComboBoxItem { Content = "Все металлы" });
                    foreach (var metal in db.MetalsCatalog.ToList())
                    {
                        cmbItemMetal.Items.Add(new ComboBoxItem
                        {
                            Content = $"{metal.MetalName} ({metal.PurityValue})",
                            Tag = metal.AssayID.ToString()
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка загрузки фильтров: " + ex.Message);
            }
        }

        private void btnFirstPage_Click(object sender, RoutedEventArgs e) =>
            LoadWarehouseItems(1, _warehousePageSize);

        private void btnPrevPage_Click(object sender, RoutedEventArgs e) =>
            LoadWarehouseItems(_warehouseCurrentPage - 1, _warehousePageSize);

        private void btnNextPage_Click(object sender, RoutedEventArgs e) =>
            LoadWarehouseItems(_warehouseCurrentPage + 1, _warehousePageSize);

        private void btnLastPage_Click(object sender, RoutedEventArgs e) =>
            LoadWarehouseItems(_warehouseTotalPages, _warehousePageSize);

        private void txtCurrentPage_PreviewTextInput(object sender, TextCompositionEventArgs e) =>
            e.Handled = !int.TryParse(e.Text, out _);

        private void txtCurrentPage_LostFocus(object sender, RoutedEventArgs e)
        {
            if (int.TryParse(txtCurrentPage.Text, out int page) && page > 0 && page <= _warehouseTotalPages)
                LoadWarehouseItems(page, _warehousePageSize);
            else
                txtCurrentPage.Text = _warehouseCurrentPage.ToString();
        }

        private void txtCurrentPage_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter && int.TryParse(txtCurrentPage.Text, out int page) && page > 0 && page <= _warehouseTotalPages)
            {
                LoadWarehouseItems(page, _warehousePageSize);
                Keyboard.ClearFocus();
            }
        }

        private void cmbPageSize_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (cmbPageSize.SelectedItem is ComboBoxItem item && int.TryParse(item.Content.ToString(), out int size))
            {
                LoadWarehouseItems(1, size);
            }
        }

        // ═══════════════════════════════════════════════════════════════
        //                    АДМИН - ПОИСК И ФИЛЬТРЫ
        // ═══════════════════════════════════════════════════════════════

        private void txtSearchUsers_TextChanged(object sender, TextChangedEventArgs e)
        {
            _userSearch = txtSearchUsers.Text.ToLower();
            LoadUsersWithFilters();
        }

        private void cmbUserRole_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (cmbUserRole.SelectedItem is ComboBoxItem item)
            {
                if (item.Content.ToString() == "Все роли")
                    _userRoleFilter = null;
                else if (item.Tag != null && int.TryParse(item.Tag.ToString(), out int roleId))
                    _userRoleFilter = roleId;
                LoadUsersWithFilters();
            }
        }

        private void btnResetUserFilters_Click(object sender, RoutedEventArgs e)
        {
            txtSearchUsers.Text = "";
            cmbUserRole.SelectedIndex = 0;
            LoadUsersWithFilters();
        }

        private void LoadUsersWithFilters()
        {
            try
            {
                using (var db = new LombardDBEntities())
                {
                    var query = db.Users.Include(u => u.UserRoles).AsQueryable();

                    if (!string.IsNullOrEmpty(_userSearch))
                    {
                        query = query.Where(u => u.FullName.ToLower().Contains(_userSearch) ||
                                                u.Login.ToLower().Contains(_userSearch));
                    }

                    if (_userRoleFilter.HasValue)
                    {
                        query = query.Where(u => u.RoleID == _userRoleFilter);
                    }

                    // Добавлена проверка на null
                    if (gridUsers != null)
                    {
                        gridUsers.ItemsSource = query.ToList();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка загрузки пользователей: " + ex.Message);
            }
        }

        private void LoadUserRoles()
        {
            try
            {
                using (var db = new LombardDBEntities())
                {
                    cmbUserRole.Items.Clear();
                    cmbUserRole.Items.Add(new ComboBoxItem { Content = "Все роли" });
                    foreach (var role in db.UserRoles.ToList())
                    {
                        cmbUserRole.Items.Add(new ComboBoxItem
                        {
                            Content = role.RoleName,
                            Tag = role.RoleID.ToString()
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка загрузки ролей: " + ex.Message);
            }
        }

        private void LoadUsers()
        {
            try
            {
                using (var db = new LombardDBEntities())
                {
                    if (gridUsers != null)
                        gridUsers.ItemsSource = db.Users.Include(u => u.UserRoles).ToList();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка загрузки пользователей: " + ex.Message);
            }
        }

        private void btnAddEmployee_Click(object sender, RoutedEventArgs e)
        {
            var win = new UserEditWindow();
            if (win.ShowDialog() == true)
            {
                using (var db = new LombardDBEntities())
                {
                    db.Users.Add(win.Employee);
                    db.SaveChanges();
                }
                LoadUsersWithFilters();
            }
        }

        private void btnEditEmployee_Click(object sender, RoutedEventArgs e)
        {
            if (gridUsers.SelectedItem is Users selected)
            {
                var win = new UserEditWindow(selected);
                if (win.ShowDialog() == true)
                {
                    using (var db = new LombardDBEntities())
                    {
                        db.Entry(selected).State = EntityState.Modified;
                        db.SaveChanges();
                    }
                    LoadUsersWithFilters();
                }
            }
        }
        //                    ОБЩИЕ МЕТОДЫ
        private void btnNewContract_Click(object sender, RoutedEventArgs e)
        {
            var win = new NewContractWindow();
            if (win.ShowDialog() == true) LoadContractsWithFilters();
        }

        private void btnFinishContract_Click(object sender, RoutedEventArgs e)
        {
            if (!(gridContracts.SelectedItem is Contracts selectedContract))
            {
                MessageBox.Show("Пожалуйста, выберите договор в таблице для закрытия.",
                                "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (selectedContract.Status == "Закрыт")
            {
                MessageBox.Show("Этот договор уже закрыт.", "Внимание",
                                MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var result = MessageBox.Show(
                $"Вы уверены, что хотите закрыть договор №{selectedContract.ContractID}?\n" +
                $"Клиент: {selectedContract.Clients.LastName}\n" +
                $"Сумма к возврату: {selectedContract.LoanAmount} ₽",
                "Закрытие договора", MessageBoxButton.YesNo, MessageBoxImage.Question);

            if (result == MessageBoxResult.Yes)
            {
                try
                {
                    using (var db = new LombardDBEntities())
                    {
                        db.Contracts.Attach(selectedContract);
                        selectedContract.Status = "Закрыт";
                        db.Entry(selectedContract).State = EntityState.Modified;
                        db.SaveChanges();
                    }
                    MessageBox.Show("Договор успешно закрыт. Залог выдан клиенту.", "Успех",
                                    MessageBoxButton.OK, MessageBoxImage.Information);
                    LoadContractsWithFilters();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка при закрытии договора: " + ex.Message);
                }
            }
        }

        private void btnSellItem_Click(object sender, RoutedEventArgs e)
        {
            if (!(gridContracts.SelectedItem is Contracts selectedContract))
            {
                MessageBox.Show("Выберите договор в таблице.", "Предупреждение",
                                MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (selectedContract.Status != "Просрочен")
            {
                MessageBox.Show(
                    "Продажа изделий доступна только для договоров со статусом «Просрочен».\n\n" +
                    "Текущий статус договора: " + selectedContract.Status,
                    "Внимание", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            try
            {
                using (var db = new LombardDBEntities())
                {
                    // Загружаем контракт вместе со всеми вложенными данными, которые могут понадобиться
                    var contractWithDetails = db.Contracts
                        .Include(c => c.Clients)
                        .Include("ContractItems.Items.ItemTypes") // Загружаем предметы и их типы
                        .Include("ContractItems.Items.MetalsCatalog") // Загружаем данные о металле
                        .FirstOrDefault(c => c.ContractID == selectedContract.ContractID);

                    if (contractWithDetails != null)
                    {
                        var saleWin = new SaleItemWindow(contractWithDetails);

                        // Окно открывается внутри блока using, или данные уже в памяти
                        if (saleWin.ShowDialog() == true)
                            LoadContractsWithFilters();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка: " + ex.Message);
            }
        }

        private void btnEditRate_Click(object sender, RoutedEventArgs e)
        {
            if (!(gridPurchaseRates.SelectedItem is PurchaseRates selected))
            {
                MessageBox.Show("Выберите строку с металлом для изменения цены.",
                                "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                using (var db = new LombardDBEntities())
                {
                    var rate = db.PurchaseRates
                        .Include(r => r.MetalsCatalog)
                        .FirstOrDefault(r => r.RateID == selected.RateID);

                    if (rate == null) return;

                    var win = new EditRate(rate) { Owner = this };
                    if (win.ShowDialog() == true)
                    {
                        db.Entry(rate).State = EntityState.Modified;
                        db.SaveChanges();

                        Logger.Write("Изменена цена металла",
                            $"{rate.MetalsCatalog?.MetalName} → {rate.PricePerGram:F2} ₽");

                        LoadDataFromDb();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при сохранении цены: " + ex.Message);
            }
        }

        private void ExportContractToWord(Contracts contract)
        {
            Word.Application wordApp = null;
            Word.Document doc = null;

            try
            {
                // 1. ПОДГОТОВКА ДАННЫХ (Eager Loading)
                using (var db = new LombardDBEntities())
                {
                    // Используем ID из переданного объекта, чтобы подтянуть свежие данные со всеми связями
                    int id = contract.ContractID;
                    contract = db.Contracts
                        .Include("Clients")
                        .Include("ContractItems.Items.ItemTypes")
                        .Include("ContractItems.Items.MetalsCatalog")
                        .FirstOrDefault(c => c.ContractID == id);
                }

                if (contract == null) return;

                string templatePath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Template.docx");

                if (!System.IO.File.Exists(templatePath))
                {
                    MessageBox.Show("Файл шаблона не найден по пути: " + templatePath);
                    return;
                }

                wordApp = new Word.Application();
                // Скрываем Word на время заполнения, чтобы не было мерцания и зависаний интерфейса
                wordApp.Visible = false;

                // ИСПРАВЛЕНИЕ: Добавляем ReadOnly: true, чтобы убрать ошибку "File In Use"
                doc = wordApp.Documents.Open(templatePath, ReadOnly: true);

                // Заполнение полей
                ReplaceBookmarkText(doc, "ContractID", contract.ContractID.ToString());
                ReplaceBookmarkText(doc, "FullName", $"{contract.Clients.LastName} {contract.Clients.FirstName} {contract.Clients.MiddleName}");
                ReplaceBookmarkText(doc, "Passport", contract.Clients.Passport);
                ReplaceBookmarkText(doc, "LoanAmount", contract.LoanAmount.ToString("N2"));
                ReplaceBookmarkText(doc, "Status", contract.Status);
                ReplaceBookmarkText(doc, "CurrentDate", DateTime.Now.ToString("dd.MM.yyyy"));
                ReplaceBookmarkText(doc, "FullNameEnd", $"{contract.Clients.LastName} {contract.Clients.FirstName} {contract.Clients.MiddleName}");

                // 2. ЗАПОЛНЕНИЕ ТАБЛИЦЫ
                FillItemsTableOptimized(doc, contract);

                // Показываем результат пользователю
                wordApp.Visible = true;
                wordApp.Activate(); // Выводим окно Word на передний план
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при работе с Word: " + ex.Message);

                // В случае ошибки закрываем всё, чтобы процессы не висели в памяти
                if (doc != null) doc.Close(false);
                if (wordApp != null) wordApp.Quit();
            }
        }

        private void FillItemsTableOptimized(Word.Document doc, Contracts contract)
        {
            if (!doc.Bookmarks.Exists("ItemsTable")) return;

            Word.Range range = doc.Bookmarks["ItemsTable"].Range;
            Word.Table table = range.Tables[1];

            int rowIndex = 2;
            int counter = 1;

            foreach (var ci in contract.ContractItems)
            {
                if (counter > 1) table.Rows.Add();

                var item = ci.Items; // Данные уже загружены благодаря .Include() выше

                table.Cell(rowIndex, 1).Range.Text = counter.ToString();
                table.Cell(rowIndex, 2).Range.Text = item?.ItemTypes?.TypeName ?? "Изделие";

                string metal = item?.MetalsCatalog?.MetalName ?? "";
                string purity = item?.MetalsCatalog?.PurityValue.ToString() ?? "";
                table.Cell(rowIndex, 3).Range.Text = $"{metal} {purity}".Trim();

                if (item != null)
                {
                    table.Cell(rowIndex, 4).Range.Text = item.Weight.ToString("F2") + " г";
                }
                else
                {
                    table.Cell(rowIndex, 4).Range.Text = "0.00 г";
                }
                table.Cell(rowIndex, 5).Range.Text = (ci.EstimatedValue ?? 0).ToString("F2") + " руб.";
                table.Cell(rowIndex, 6).Range.Text = contract.LoanAmount.ToString("F2") + " руб.";

                rowIndex++;
                counter++;
            }
        }

        private void ReplaceBookmarkText(Word.Document doc, string bookmarkName, string text)
        {
            if (doc.Bookmarks.Exists(bookmarkName))
            {
                Word.Range range = doc.Bookmarks[bookmarkName].Range;
                range.Text = text;
                // Чтобы закладка не удалялась после вставки текста (полезно для повторного использования)
                doc.Bookmarks.Add(bookmarkName, range);
            }
        }

        private void ExportStatsToExcel()
        {
            try
            {
                using (var db = new LombardDBEntities())
                {
                    var data = db.Contracts.Include(c => c.Clients).ToList();

                    var excelApp = new Excel.Application();
                    excelApp.Visible = true;
                    Excel.Workbook workbook = excelApp.Workbooks.Add();
                    Excel._Worksheet worksheet = workbook.ActiveSheet;

                    worksheet.Cells[1, 1] = "№ Контракта";
                    worksheet.Cells[1, 2] = "Клиент";
                    worksheet.Cells[1, 3] = "Сумма займа";
                    worksheet.Cells[1, 4] = "Статус";

                    int row = 2;
                    foreach (var item in data)
                    {
                        worksheet.Cells[row, 1] = item.ContractID;
                        worksheet.Cells[row, 2] = item.Clients.LastName;
                        worksheet.Cells[row, 3] = item.LoanAmount;
                        worksheet.Cells[row, 4] = item.Status;
                        row++;
                    }

                    Excel.Range headerRange = worksheet.get_Range("A1", "D1");
                    headerRange.Font.Bold = true;
                    headerRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                    worksheet.Cells[row, 2] = "ИТОГО:";
                    worksheet.Cells[row, 3].Formula = $"=SUM(C2:C{row - 1})";
                    worksheet.Cells[row, 3].Font.Bold = true;

                    // Добавляем диаграмму
                    if (data.Count > 0)
                    {
                        // Создаем сводные данные для диаграммы по статусам
                        var statusData = data.GroupBy(c => c.Status)
                                            .Select(g => new { Status = g.Key, Count = g.Count(), Sum = g.Sum(c => (decimal?)c.LoanAmount) ?? 0 })
                                            .ToList();

                        // Добавляем лист для диаграммы
                        Excel._Worksheet chartSheet = (Excel._Worksheet)workbook.Worksheets.Add();
                        chartSheet.Name = "Диаграмма статусов";

                        // Заголовки
                        chartSheet.Cells[1, 1] = "Статус";
                        chartSheet.Cells[1, 2] = "Количество";
                        chartSheet.Cells[1, 3] = "Сумма (₽)";

                        int chartRow = 2;
                        foreach (var status in statusData)
                        {
                            chartSheet.Cells[chartRow, 1] = status.Status;
                            chartSheet.Cells[chartRow, 2] = status.Count;
                            chartSheet.Cells[chartRow, 3] = status.Sum;
                            chartRow++;
                        }

                        // Создаем диаграмму
                        Excel.ChartObjects chartObjects = (Excel.ChartObjects)chartSheet.ChartObjects();
                        Excel.ChartObject chartObj = chartObjects.Add(50, 150, 400, 300);
                        Excel.Chart chart = chartObj.Chart;

                        // Устанавливаем источник данных
                        Excel.Range dataRange = chartSheet.get_Range("A1", $"C{chartRow - 1}");
                        chart.SetSourceData(dataRange);

                        // Тип диаграммы - гистограмма
                        chart.ChartType = Excel.XlChartType.xlColumnClustered;
                        chart.HasTitle = true;
                        chart.ChartTitle.Text = "Распределение контрактов по статусам";

                        // Настройка осей
                        chart.Axes(Excel.XlAxisType.xlCategory).HasTitle = true;
                        chart.Axes(Excel.XlAxisType.xlCategory).AxisTitle.Text = "Статус";
                        chart.Axes(Excel.XlAxisType.xlValue).HasTitle = true;
                        chart.Axes(Excel.XlAxisType.xlValue).AxisTitle.Text = "Количество";

                        // Показываем легенду
                        chart.HasLegend = true;
                        chart.Legend.Position = Excel.XlLegendPosition.xlLegendPositionBottom;

                        worksheet.Columns.AutoFit();
                        MessageBox.Show("Отчет Excel сформирован!\nДиаграмма добавлена на отдельный лист.", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    else
                    {
                        worksheet.Columns.AutoFit();
                        MessageBox.Show("Отчет Excel сформирован!");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при работе с Excel: " + ex.Message);
            }
        }

        private void btnExportExcel_Click(object sender, RoutedEventArgs e) => ExportStatsToExcel();

        private void btnPrintContract_Click(object sender, RoutedEventArgs e)
        {
            if (gridContracts.SelectedItem is Contracts selected)
            {
                ExportContractToWord(selected);
            }
            else
            {
                MessageBox.Show("Выберите контракт для печати");
            }
        }

        private void btnOpenLogs_Click(object sender, RoutedEventArgs e)
        {
            var logsWin = new LogViwerWindow();
            logsWin.ShowDialog();
        }

        private void Logout_Click(object sender, RoutedEventArgs e)
        {
            Logger.Write("Выход из системы", "");
            new LoginWindow().Show();
            this.Close();
        }

        private void MainContent_SelectionChanged(object sender, SelectionChangedEventArgs e) { }
    }
}