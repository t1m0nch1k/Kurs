using System;
using System.Linq;
using System.Windows;
using LiveCharts;
using LiveCharts.Wpf;

namespace Kurs
{
    public partial class ContractChartWindow : Window
    {
        public ContractChartWindow()
        {
            InitializeComponent();
            LoadChartData();
        }

        private void LoadChartData()
        {
            try
            {
                using (var db = new LombardDBEntities())
                {
                    var allContracts = db.Contracts.ToList();

                    // Подсчет по статусам
                    var activeCount = allContracts.Count(c => c.Status == "Активен");
                    var overdueCount = allContracts.Count(c => c.Status == "Просрочен");
                    var closedCount = allContracts.Count(c => c.Status == "Закрыт");

                    // Обновление карточек со счетчиками
                    txtActiveCount.Text = activeCount.ToString();
                    txtOverdueCount.Text = overdueCount.ToString();
                    txtClosedCount.Text = closedCount.ToString();

                    // Общая статистика
                    txtTotalContracts.Text = allContracts.Count.ToString();
                    
                    decimal totalAmount = allContracts.Sum(c => (decimal?)c.LoanAmount) ?? 0;
                    decimal activeAmount = allContracts.Where(c => c.Status == "Активен").Sum(c => (decimal?)c.LoanAmount) ?? 0;
                    decimal overdueAmount = allContracts.Where(c => c.Status == "Просрочен").Sum(c => (decimal?)c.LoanAmount) ?? 0;

                    txtTotalAmount.Text = totalAmount.ToString("N0") + " ₽";
                    txtActiveAmount.Text = activeAmount.ToString("N0") + " ₽";
                    txtOverdueAmount.Text = overdueAmount.ToString("N0") + " ₽";

                    // Создание диаграммы
                    var seriesCollection = new SeriesCollection
                    {
                        new ColumnSeries
                        {
                            Title = "Количество контрактов",
                            Values = new ChartValues<int> { activeCount, overdueCount, closedCount },
                            Fill = System.Windows.Media.Brushes.SteelBlue
                        }
                    };

                    chartStatus.Series = seriesCollection;
                    chartStatus.AxisX[0].Labels = new[] { "Активные", "Просроченные", "Закрытые" };
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка загрузки данных диаграммы: " + ex.Message);
            }
        }

        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
