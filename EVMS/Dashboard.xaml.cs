using EVMS.Service;
using LiveChartsCore.Kernel;
using MetroGaugeSoft;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using static EVMS.Login_Page;

namespace EVMS
{
    public partial class Dashboard : UserControl, INotifyPropertyChanged
    {
        // ===== YOUR OLD DEMO VALUES (KEPT – NOT USED NOW) =====
        private double[] _values = new double[] { 25, 8, 32, 5, 28, 12, 41, 3 };
        public double[] Values
        {
            get => _values;
            set { _values = value; OnPropertyChanged(); }
        }

        public Func<double, string> LabelFormatter { get; set; } = value => value.ToString("N0");

        private double _vanesa = 30, _charles = 50, _ana = 70;
        public double Vanesa { get => _vanesa; set { _vanesa = value; OnPropertyChanged(); } }
        public double Charles { get => _charles; set { _charles = value; OnPropertyChanged(); } }
        public double Ana { get => _ana; set { _ana = value; OnPropertyChanged(); } }

        // ================= X AXIS =================
        private string[] _periodLabels = Array.Empty<string>();
        public string[] PeriodLabels
        {
            get => _periodLabels;
            set { _periodLabels = value; OnPropertyChanged(); }
        }

        // ================= LINE DATA =================
        private double[] _okValues = Array.Empty<double>();
        public double[] OkValues
        {
            get => _okValues;
            set { _okValues = value; OnPropertyChanged(); }
        }

        private double[] _ngValues = Array.Empty<double>();
        public double[] NgValues
        {
            get => _ngValues;
            set { _ngValues = value; OnPropertyChanged(); }
        }



        public Func<ChartPoint, string> LabelFormatter1 { get; set; } =
            point => $"{point.Coordinate.PrimaryValue:F1}% ({point.Context.Series.Name})";

        public Grid MainContentGrid { get; set; }
        public event Action<string> StatusMessageChanged;
        public event Action<string> DashboardStatusMessageChanged;

        // ================= SERVICE =================
        private readonly DataStorageService _dataStorageService;
        private string _selectedPartNo = "All";
        private string _selectedPeriod = "month";

        // ================= CARDS =================
        private int _totalProduction;
        public int TotalProduction
        {
            get => _totalProduction;
            set { _totalProduction = value; OnPropertyChanged(); }
        }

        private int _okParts;
        public int OKParts
        {
            get => _okParts;
            set { _okParts = value; OnPropertyChanged(); }
        }

        private int _ngParts;
        public int NGParts
        {
            get => _ngParts;
            set { _ngParts = value; OnPropertyChanged(); }
        }

        // ================= NOTIFY =================
        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string name = null) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));

        // ================= CONSTRUCTOR =================
        public Dashboard()
        {
            _dataStorageService = new DataStorageService();

            InitializeComponent();
            DataContext = this;
            Loaded += Dashboard_Loaded;
        }

        // ================= LOAD =================
        private async void Dashboard_Loaded(object sender, RoutedEventArgs e)
        {
            await LoadDashboardData();
            SetButtonAccessByRole();
        }

        private async void PeriodCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (sender is ComboBox combo && combo.SelectedItem is ComboBoxItem item)
            {
                _selectedPeriod = item.Content.ToString();
                await LoadDashboardData();
            }
        }


        private async Task LoadDashboardData()
        {
            try
            {
                var activeParts = _dataStorageService.GetActiveParts();
                if (activeParts == null || activeParts.Count == 0)
                {
                    UpdateStatus("No active part.");
                    return;
                }

                string partNumber = activeParts[0]?.Para_No ?? "";

                DateTime fromDate;
                DateTime toDate = DateTime.Today;

                // Decide range
                switch (_selectedPeriod)
                {
                    case "Year":
                        fromDate = new DateTime(DateTime.Today.Year - 5, 1, 1);
                        break;

                    case "Month":
                        fromDate = DateTime.Today.AddMonths(-12);
                        break;

                    default: // Day
                        fromDate = DateTime.Today.AddDays(-30);
                        break;
                }

                var stats = _dataStorageService.GetProductionCountsForActivePart(
                    partNo: partNumber,
                    fromDate: fromDate,
                    toDate: toDate,
                    period: _selectedPeriod);

                if (stats == null || stats.Count == 0)
                {
                    UpdateStatus("No data found.");
                    OkValues = Array.Empty<double>();
                    NgValues = Array.Empty<double>();
                    PeriodLabels = Array.Empty<string>();
                    return;
                }

                // ===============================
                // STAT CARDS
                // ===============================
                TotalProduction = stats.Sum(x => x.TotalCount);
                OKParts = stats.Sum(x => x.OKCount);
                NGParts = stats.Sum(x => x.NGCount);

                OnPropertyChanged(nameof(TotalProduction));
                OnPropertyChanged(nameof(OKParts));
                OnPropertyChanged(nameof(NGParts));

                // ===============================
                // CHART VALUES
                // ===============================
                OkValues = stats.Select(x => (double)x.OKCount).ToArray();
                NgValues = stats.Select(x => (double)x.NGCount).ToArray();

                // ===============================
                // X AXIS LABELS
                // ===============================
                if (_selectedPeriod.Equals("Month", StringComparison.OrdinalIgnoreCase))
                    PeriodLabels = stats.Select(x => x.MonthGroup).ToArray();

                else if (_selectedPeriod.Equals("Year", StringComparison.OrdinalIgnoreCase))
                    PeriodLabels = stats.Select(x => x.YearGroup.ToString()).ToArray();

                else
                    PeriodLabels = stats
                        .Select(x => x.DayGroup.HasValue ? x.DayGroup.Value.ToString("dd") : "")
                        .ToArray();


                // ===============================
                // GAUGE %
                // ===============================
                int total = TotalProduction;
                Charles = total > 0 ? (double)OKParts / total * 100 : 0;
                Ana = total > 0 ? (double)NGParts / total * 100 : 0;

                UpdateStatus($"Dashboard: {TotalProduction} parts ({OKParts} OK, {NGParts} NG)");
            }
            catch (Exception ex)
            {
                UpdateStatus($"Data error: {ex.Message}");

                // demo fallback
                OkValues = new double[] { 10, 20, 15, 30 };
                NgValues = new double[] { 2, 5, 3, 6 };
                PeriodLabels = new string[] { "1", "2", "3", "4" };

                Ana = 70;
                Charles = 30;
            }
        }


        // ================= REFRESH =================
        public async void RefreshDashboard(string partNo = "All", string period = "month")
        {
            _selectedPartNo = partNo;
            _selectedPeriod = period.ToLower();
            await LoadDashboardData();
        }

        // ================= YOUR OLD METHODS (UNCHANGED) =================
        private void SetButtonAccessByRole()
        {
            string userType = SessionManager.UserType;
            if (string.Equals(userType, "Admin", StringComparison.OrdinalIgnoreCase))
            {
                SetAllButtonsEnabled(true);
            }
            else
            {
                SetButtonsEnabled(new List<string>
                {
                    "Login/Setup", "Part Manager", "Settings", "Calculation"
                }, false);
            }
        }

        private void UpdateStatus(string message)
        {
            StatusMessageChanged?.Invoke(message);
            DashboardStatusMessageChanged?.Invoke(message);
        }

        private void SetAllButtonsEnabled(bool enabled)
        {
            foreach (var btn in FindButtons(this))
                btn.IsEnabled = enabled;
        }

        private void SetButtonsEnabled(List<string> buttonContents, bool enabled)
        {
            foreach (var btn in FindButtons(this))
            {
                if (btn.Content is string content && buttonContents.Contains(content))
                    btn.IsEnabled = enabled;
            }
        }

        private IEnumerable<Button> FindButtons(DependencyObject parent)
        {
            if (parent == null) yield break;
            int count = VisualTreeHelper.GetChildrenCount(parent);
            for (int i = 0; i < count; i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);
                if (child is Button btn) yield return btn;
                foreach (var childOfChild in FindButtons(child)) yield return childOfChild;
            }
        }

        // ================= NAVIGATION (UNCHANGED) =================
        private void DashboardCard_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button button && button.Content is string key)
            {
                UserControl pageToShow = null;

                switch (key)
                {
                    case "Measurement":
                        var entryPage = new EntryPage();
                        entryPage.StartClicked += EntryPage_StartClicked;
                        var currentWindowForEntry = Window.GetWindow(this);
                        entryPage.MainContentGrid = currentWindowForEntry?.FindName("MainContentGrid") as Grid;
                        pageToShow = entryPage;
                        break;

                    case "Probe Setup": pageToShow = new ProbeSetupPage(); break;
                    case "Master Reading": pageToShow = new MasterReadingPage(); break;
                    case "Login/Setup": pageToShow = new AdminControlePage(); break;
                    case "Report Graph": pageToShow = new Report_GraphPage(); break;
                    case "Report View": pageToShow = new Report_View_Page(); break;
                    case "Repeatability": pageToShow = new Repeatbilty_Page(); break;
                    case "IO Control": pageToShow = new IO_Controle_page(); break;
                    case "Settings": pageToShow = new SettingsPage(); break;
                    case "Probes": pageToShow = new ProbeInstallPage(); break;
                    case "Part Config": pageToShow = new PartConfig(); break;
                    case "Part Manager": pageToShow = new Part_Manager(); break;
                    case "Calculation": pageToShow = new Calculation_Modification(); break;
                    case "RnR Report": pageToShow = new RnR_Report_Page(); break;
                    case "Maintenance": pageToShow = new MaintenanceControl(); break;


                }

                if (pageToShow != null)
                {
                    var currentWindow = Window.GetWindow(this);
                    var mainContentGrid = currentWindow?.FindName("MainContentGrid") as Grid;

                    if (mainContentGrid != null)
                    {
                        mainContentGrid.Children.Clear();
                        pageToShow.HorizontalAlignment = HorizontalAlignment.Stretch;
                        pageToShow.VerticalAlignment = VerticalAlignment.Stretch;
                        mainContentGrid.Children.Add(pageToShow);
                    }
                    else
                    {
                        MessageBox.Show("MainContentGrid not found in MainWindow.");
                    }
                }
            }
        }

        private void EntryPage_StartClicked(object sender, StartClickedEventArgs e)
        {
            Grid mainContentGrid = null;

            if (sender is EntryPage entryPage && entryPage.MainContentGrid != null)
                mainContentGrid = entryPage.MainContentGrid;
            else
                mainContentGrid = Window.GetWindow(this)?.FindName("MainContentGrid") as Grid;

            if (mainContentGrid != null)
            {
                mainContentGrid.Children.Clear();

                var resultPage = new ResultPage(e.Model, e.LotNo, e.UserId)
                {
                    HorizontalAlignment = HorizontalAlignment.Stretch,
                    VerticalAlignment = VerticalAlignment.Stretch
                };

                mainContentGrid.Children.Add(resultPage);

                resultPage.StatusMessageChanged += (message) =>
                {
                    DashboardStatusMessageChanged?.Invoke(message);
                };
            }
            else
            {
                MessageBox.Show("MainContentGrid not found in MainWindow. Cannot open ResultPage.");
            }
        }
    }
}
