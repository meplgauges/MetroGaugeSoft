using EVMS.Service;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace EVMS
{
    public partial class Repeatbilty_Page : UserControl, INotifyPropertyChanged
    {
        private readonly DataStorageService _dataService;

        public ObservableCollection<string> ActiveParts { get; set; } = new();
        public ObservableCollection<string> ParametersOptions { get; set; } = new();
        public ObservableCollection<RepeatReadingItem> RepeatabilityItems { get; set; } = new();

        private string _selectedPartNo;
        public string SelectedPartNo
        {
            get => _selectedPartNo;
            set
            {
                if (_selectedPartNo != value)
                {
                    _selectedPartNo = value;
                    OnPropertyChanged(nameof(SelectedPartNo));
                    LoadParametersForPart();
                }
            }
        }

        private string _selectedParameter;
        public string SelectedParameter
        {
            get => _selectedParameter;
            set
            {
                if (_selectedParameter != value)
                {
                    _selectedParameter = value;
                    OnPropertyChanged(nameof(SelectedParameter));

                    if (!string.IsNullOrEmpty(_selectedPartNo) && !string.IsNullOrEmpty(_selectedParameter))
                    {
                        _ = LoadReadingsAsync();
                    }
                }
            }
        }

        public ObservableCollection<int> RecordsCountOptions { get; set; } = new();
        private int _selectedRecordsCount = 30;   // default 30

        public int SelectedRecordsCount
        {
            get => _selectedRecordsCount;
            set
            {
                if (_selectedRecordsCount != value)
                {
                    _selectedRecordsCount = value;
                    OnPropertyChanged(nameof(SelectedRecordsCount));

                    // Reload when user changes count and parameter is selected
                    if (!string.IsNullOrEmpty(SelectedPartNo) &&
                        !string.IsNullOrEmpty(SelectedParameter))
                    {
                        _ = LoadReadingsAsync();
                    }
                }
            }
        }

        private SummaryStats _summary = new();
        public SummaryStats Summary
        {
            get => _summary;
            set { _summary = value; OnPropertyChanged(nameof(Summary)); }
        }

        public Repeatbilty_Page()
        {
            InitializeComponent();
            Keyboard.Focus(this);
            this.Focus();

            this.PreviewKeyDown += Repeatbilty_Page_PreviewKeyDown;

            _dataService = new DataStorageService();
            DataContext = this;

            // ✅ Set fixed options for number of records
            RecordsCountOptions.Clear();
            foreach (var n in new[] { 5, 10, 15, 20, 25, 30 })
                RecordsCountOptions.Add(n);

            SelectedRecordsCount = 30; // default

            LoadActiveParts();
        }


        private void LoadActiveParts()
        {
            ActiveParts.Clear();
            var parts = _dataService.GetActiveParts();

            foreach (var p in parts)
                ActiveParts.Add(p.Para_No);

            // ✅ Auto-pick first active part
            if (ActiveParts.Any())
                SelectedPartNo = ActiveParts[0];   // This will call LoadParametersForPart()
        }


        private void LoadParametersForPart()
        {
            ParametersOptions.Clear();
            if (string.IsNullOrEmpty(SelectedPartNo))
                return;

            var partConfig = _dataService.GetPartConfigByPartNumber(SelectedPartNo);
            foreach (var cfg in partConfig)
                ParametersOptions.Add(cfg.Parameter);
        }

        private async Task LoadReadingsAsync()
        {
            try
            {
                RepeatabilityItems.Clear();
                Summary = new SummaryStats();

                if (string.IsNullOrEmpty(SelectedPartNo) || string.IsNullOrEmpty(SelectedParameter))
                    return;

                var readings = await _dataService.GetMeasurementReadingsAsync(
                    SelectedPartNo, null, null, null, null);

                if (readings == null || !readings.Any())
                    return;

                string columnName = MapParameterToProperty(SelectedParameter);

                int count = SelectedRecordsCount > 0 ? SelectedRecordsCount : 30;

                var lastN = readings.OrderByDescending(r => r.MeasurementDate)
                                    .Take(count)
                                    .ToList();

                var values = lastN
                    .Select(r =>
                    {
                        var prop = r.GetType().GetProperty(columnName);
                        if (prop == null) return (Reading: 0.0, Row: r);
                        var val = prop.GetValue(r);
                        return (Reading: Convert.ToDouble(val ?? 0), Row: r);
                    })
                    .Where(v => v.Reading != 0)
                    .ToList();


                int serialNo = 1;
                foreach (var v in values)
                {
                    RepeatabilityItems.Add(new RepeatReadingItem
                    {
                        SerialNo = serialNo++,
                        LotNo = v.Row.LotNo,
                        Operator = v.Row.Operator_ID,
                        Reading = v.Reading,
                        DateTime = v.Row.MeasurementDate.ToString("MM/dd/yyyy HH:mm:ss")
                    });
                }

                if (values.Any())
                {
                    var partCfg = _dataService.GetPartConfigByPartNumber(SelectedPartNo)
                                              .FirstOrDefault(p => p.Parameter == SelectedParameter);
                    double min = values.Min(x => x.Reading);
                    double max = values.Max(x => x.Reading);
                    double variation = max - min;

                    Summary = new SummaryStats
                    {
                        USL = partCfg?.Nominal + (partCfg?.RTolPlus ?? 0) ?? 0,
                        LSL = partCfg?.Nominal - (partCfg?.RTolMinus ?? 0) ?? 0,
                        Min = min,
                        Max = max,
                        Variation = variation
                    };
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading readings: {ex.Message}", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private string MapParameterToProperty(string parameter) => parameter switch
        {

            "OD1" => nameof(MeasurementReading.StepOd1),
            "RN1" => nameof(MeasurementReading.StepRunout1),
            "OD2" => nameof(MeasurementReading.Od1),
            "RN2" => nameof(MeasurementReading.Rn1),
            "OD3" => nameof(MeasurementReading.Od2),
            "RN3" => nameof(MeasurementReading.Rn2),
            "OD4" => nameof(MeasurementReading.Od3),
            "RN4" => nameof(MeasurementReading.Rn3),
            "OD5" => nameof(MeasurementReading.StepOd2),
            "RN5" => nameof(MeasurementReading.StepRunout2),
            "ID-1" => nameof(MeasurementReading.Id1),
            "RN6" => nameof(MeasurementReading.Rn4),
            "ID-2" => nameof(MeasurementReading.Id2),
            "RN7" => nameof(MeasurementReading.Rn5),
            "OL" => nameof(MeasurementReading.Ol),
            _ => parameter.Replace(" ", "")
        };

        private void PartNo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (DataContext is Repeatbilty_Page vm)
                vm.SelectedPartNo = (sender as ComboBox)?.SelectedItem?.ToString();
        }

        private void Repeatbilty_Page_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                HandleEscKeyAction();
                e.Handled = true;
            }
        }

        private void HandleEscKeyAction()
        {
            try
            {
                Window currentWindow = Window.GetWindow(this);
                if (currentWindow != null)
                {
                    var mainContentGrid = currentWindow.FindName("MainContentGrid") as Grid;
                    if (mainContentGrid != null)
                    {
                        mainContentGrid.Children.Clear();

                        var homePage = new Dashboard
                        {
                            HorizontalAlignment = HorizontalAlignment.Stretch,
                            VerticalAlignment = VerticalAlignment.Stretch
                        };

                        mainContentGrid.Children.Add(homePage);
                        mainContentGrid.UpdateLayout();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error navigating to home page: {ex.Message}", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string name) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    public class RepeatReadingItem
    {
        public int SerialNo { get; set; }
        public string LotNo { get; set; }
        public string Operator { get; set; }
        public double Reading { get; set; }
        public string DateTime { get; set; }
    }

    public class SummaryStats : INotifyPropertyChanged
    {
        private double _usl, _lsl, _min, _max, _variation;

        public double USL { get => _usl; set { _usl = value; OnPropertyChanged(nameof(USL)); } }
        public double LSL { get => _lsl; set { _lsl = value; OnPropertyChanged(nameof(LSL)); } }
        public double Min { get => _min; set { _min = value; OnPropertyChanged(nameof(Min)); } }
        public double Max { get => _max; set { _max = value; OnPropertyChanged(nameof(Max)); } }
        public double Variation { get => _variation; set { _variation = value; OnPropertyChanged(nameof(Variation)); } }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string? name) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}
