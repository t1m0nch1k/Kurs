using System;
using System.Linq;
using System.Windows;
using System.Windows.Media.Imaging;
using System.IO;

namespace Kurs
{
    public partial class ItemDetailWindow : Window
    {
        private Items _item;

        public ItemDetailWindow(Items item)
        {
            InitializeComponent();
            _item = item;
            LoadItemDetails();
        }

        private void LoadItemDetails()
        {
            if (_item == null) return;

            // Загрузка фотографии
            if (_item.ItemPhoto != null && _item.ItemPhoto.Length > 0)
            {
                var bitmap = new BitmapImage();
                bitmap.BeginInit();
                bitmap.StreamSource = new MemoryStream(_item.ItemPhoto);
                bitmap.CacheOption = BitmapCacheOption.OnLoad;
                bitmap.EndInit();
                bitmap.Freeze();
                
                imgItemPhoto.Source = bitmap;
                imgItemPhoto.Visibility = Visibility.Visible;
                txtPhotoPlaceholder.Visibility = Visibility.Collapsed;
            }
            else
            {
                imgItemPhoto.Source = null;
                imgItemPhoto.Visibility = Visibility.Collapsed;
                txtPhotoPlaceholder.Visibility = Visibility.Visible;
            }

            // Основная информация
            txtItemType.Text = _item.ItemTypes?.TypeName ?? "Не указано";
            txtMetalType.Text = _item.MetalsCatalog?.MetalName ?? "Не указано";
            txtPurity.Text = _item.MetalsCatalog?.PurityValue.ToString() ?? "Не указано";
            txtWeight.Text = $"{_item.Weight:F2} г";

            // Описание
            txtDescription.Text = _item.ItemDescription ?? "Без описания";

            // Финансовая информация
            txtEstimatedValue.Text = $"{_item.EstimatedValue:F0} ₽";
            
            // Информация о договоре (если есть)
            using (var db = new LombardDBEntities())
            {
                var contractItem = db.ContractItems.FirstOrDefault(ci => ci.ItemID == _item.ItemID);
                if (contractItem != null)
                {
                    var contract = db.Contracts.FirstOrDefault(c => c.ContractID == contractItem.ContractID);
                    if (contract != null)
                    {
                        txtContractNumber.Text = contract.ContractID.ToString();
                        txtStatus.Text = contract.Status;
                    }
                    else
                    {
                        txtContractNumber.Text = "Не привязано";
                        txtStatus.Text = "На складе";
                    }
                }
                else
                {
                    txtContractNumber.Text = "Не привязано";
                    txtStatus.Text = "На складе";
                }
            }

            // Дополнительная информация
            txtAddDate.Text = _item.AddDate.ToString("dd.MM.yyyy HH:mm");
            txtItemID.Text = _item.ItemID.ToString();
        }

        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
