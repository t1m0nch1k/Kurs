using System;
using System.Data.SqlClient;
using System.Windows;

namespace Kurs
{
    public partial class EditRate : Window
    {
        public PurchaseRates Rate { get; private set; }

        public EditRate(PurchaseRates rate)
        {
            InitializeComponent();
            Rate = rate;
            txtMetal.Text = $"{rate.MetalsCatalog?.MetalName} — проба {rate.MetalsCatalog?.PurityValue}";
            txtPrice.Text = rate.PricePerGram.ToString("F2");
        }

        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            pnlError.Visibility = Visibility.Collapsed;

            if (!decimal.TryParse(
                    txtPrice.Text.Replace(',', '.'),
                    System.Globalization.NumberStyles.Any,
                    System.Globalization.CultureInfo.InvariantCulture,
                    out decimal price) || price <= 0)
            {
                txtError.Text = "Введите корректную цену (больше 0).";
                pnlError.Visibility = Visibility.Visible;
                return;
            }

            Rate.PricePerGram = price;
            Rate.UpdateDate = DateTime.Now;
            DialogResult = true;
        }
    }
}