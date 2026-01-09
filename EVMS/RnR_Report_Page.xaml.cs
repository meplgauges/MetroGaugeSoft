using ClosedXML.Excel;
using EVMS.Service;
using Microsoft.Data.SqlClient;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;


namespace EVMS
{
    public partial class RnR_Report_Page : UserControl, INotifyPropertyChanged
    {
        private readonly DataStorageService _dataService;
        private readonly string connectionString;


        // FILTER SOURCES
        public ObservableCollection<string> ActiveParts { get; } = new();
        public ObservableCollection<string> LotNumbers { get; } = new();
        public ObservableCollection<string> ParametersOptions { get; } = new();
        public ObservableCollection<string> Operators { get; } = new();

        // GRID DATA
        public ObservableCollection<RnRGridRow> Appraiser1Data { get; } = new();
        public ObservableCollection<RnRGridRow> Appraiser2Data { get; } = new();
        public ObservableCollection<RnRGridRow> Appraiser3Data { get; } = new();

        // SUMMARY BINDINGS (displayed values)
        private double _repeatability;
        public double Repeatability { get => _repeatability; set { _repeatability = value; OnPropertyChanged(nameof(Repeatability)); } }

        private double _reproducibility;
        public double Reproducibility { get => _reproducibility; set { _reproducibility = value; OnPropertyChanged(nameof(Reproducibility)); } }

        private double _partVariation;
        public double PartVariation { get => _partVariation; set { _partVariation = value; OnPropertyChanged(nameof(PartVariation)); } }

        private double _totalPVPercent;
        public double TotalPVPercent { get => _totalPVPercent; set { _totalPVPercent = value; OnPropertyChanged(nameof(TotalPVPercent)); } }

        private double _totalTolerance;
        public double TotalTolerance { get => _totalTolerance; set { _totalTolerance = value; OnPropertyChanged(nameof(TotalTolerance)); } }

        private double _grrPercentage;
        public double GRRPercentage { get => _grrPercentage; set { _grrPercentage = value; OnPropertyChanged(nameof(GRRPercentage)); } }

        private string _grrConclusion;
        public string GRRConclusion { get => _grrConclusion; set { _grrConclusion = value; OnPropertyChanged(nameof(GRRConclusion)); } }

        // Note: Ndc kept as double for precision. Format in XAML as needed.
        private double _ndc;
        public double Ndc { get => _ndc; set { _ndc = value; OnPropertyChanged(nameof(Ndc)); } }

        // FILTER SELECTIONS
        private string _selectedPartNo;
        public string SelectedPartNo
        {
            get => _selectedPartNo;
            set { _selectedPartNo = value; OnPropertyChanged(nameof(SelectedPartNo)); _ = OnSelectedPartNoChangedAsync(); }
        }

        private string _selectedLotNo;
        public string SelectedLotNo { get => _selectedLotNo; set { _selectedLotNo = value; OnPropertyChanged(nameof(SelectedLotNo)); } }

        private string _selectedParameter;
        public string SelectedParameter { get => _selectedParameter; set { _selectedParameter = value; OnPropertyChanged(nameof(SelectedParameter)); } }

        private string _selectedOperator;
        public string SelectedOperator { get => _selectedOperator; set { _selectedOperator = value; OnPropertyChanged(nameof(SelectedOperator)); } }

        private DateTime? _selectedDateTimeFrom = DateTime.Now.AddDays(-7);
        public DateTime? SelectedDateTimeFrom
        {
            get => _selectedDateTimeFrom;
            set { _selectedDateTimeFrom = value; OnPropertyChanged(nameof(SelectedDateTimeFrom)); _ = ReloadLotAsync(); }
        }

        private DateTime? _selectedDateTimeTo = DateTime.Now;
        public DateTime? SelectedDateTimeTo
        {
            get => _selectedDateTimeTo;
            set { _selectedDateTimeTo = value; OnPropertyChanged(nameof(SelectedDateTimeTo)); _ = ReloadLotAsync(); }
        }


        // R&R method selector
        public enum RrMethod
        {
            PV,
            Tolerance
        }

        private RrMethod _selectedMethod;
        public RrMethod SelectedMethod
        {
            get => _selectedMethod;
            set
            {
                _selectedMethod = value;
                OnPropertyChanged(nameof(SelectedMethod));
                UpdateDisplayedResults();
            }
        }

        // Store both method results so user can toggle
        public GaugeRrResult PvMethodResult { get; private set; }
        public GaugeRrResult ToleranceMethodResult { get; private set; }

        private void UpdateDisplayedResults()
        {
            var active = SelectedMethod == RrMethod.PV ? PvMethodResult : ToleranceMethodResult;

            if (active == null)
            {
                ResetSummary();
                return;
            }

            Repeatability = active.Ev;
            Reproducibility = active.Av;
            PartVariation = active.Pv;
            TotalTolerance = active.Tv;
            Ndc = active.Ndc;
            TotalPVPercent = active.PercentPv;

            // Map GRR percentage and conclusion based on method
            if (SelectedMethod == RrMethod.PV)
            {
                GRRPercentage = active.PercentGrr_Pv;
                GRRConclusion = active.ConclusionTolerance; // <-- Use this
            }
            else // Tolerance method
            {
                TotalPVPercent = active.PercentGrr_Pv;
                GRRPercentage = active.PercentGrrTolerance;
                GRRConclusion = active.ConclusionTolerance;
            }
        }


        // PARAMETER → COLUMN MAP
        private static readonly Dictionary<string, string> ParameterToColumn =
     new(StringComparer.OrdinalIgnoreCase)
     {
         ["STEP OD1"] = nameof(MeasurementReading.StepOd1),
         ["STEP RUNOUT-1"] = nameof(MeasurementReading.StepRunout1),
         ["OD-1"] = nameof(MeasurementReading.Od1),
         ["RN-1"] = nameof(MeasurementReading.Rn1),
         ["OD-2"] = nameof(MeasurementReading.Od2),
         ["RN-2"] = nameof(MeasurementReading.Rn2),
         ["OD-3"] = nameof(MeasurementReading.Od3),
         ["RN-3"] = nameof(MeasurementReading.Rn3),
         ["STEP OD2"] = nameof(MeasurementReading.StepOd2),
         ["STEP RUNOUT-2"] = nameof(MeasurementReading.StepRunout2),
         ["ID-1"] = nameof(MeasurementReading.Id1),
         ["RN-4"] = nameof(MeasurementReading.Rn4),
         ["ID-2"] = nameof(MeasurementReading.Id2),
         ["RN-5"] = nameof(MeasurementReading.Rn5),
         ["OL"] = nameof(MeasurementReading.Ol)
     };

        public ICommand SubmitCommand { get; }
        public ICommand CloseCommand { get; }
        public ICommand SelectPVMethodCommand { get; }
        public ICommand SelectToleranceMethodCommand { get; }
        public ICommand RecalculateCommand { get; }
        public ICommand ExportToExcelCommand { get; }


        public event PropertyChangedEventHandler PropertyChanged;



        public RnR_Report_Page()
        {
            InitializeComponent();
            _dataService = new DataStorageService();
            connectionString = ConfigurationManager.ConnectionStrings["EVMSDb"].ConnectionString;

            SubmitCommand = new RelayCommand(async _ => await LoadRnRDataAsync(), _ => CanSubmit());
            CloseCommand = new RelayCommand(_ => ClosePage());

            SelectPVMethodCommand = new RelayCommand(_ => SelectedMethod = RrMethod.PV);
            SelectToleranceMethodCommand = new RelayCommand(_ => SelectedMethod = RrMethod.Tolerance);
            RecalculateCommand = new RelayCommand(_ => RecalculateFromGrids());
            ExportToExcelCommand = new RelayCommand(_ => ExportToExcel());

            LoadActiveParts();
            LoadOperators();
            LoadCompanyInfo();

            // default selection
            SelectedMethod = RrMethod.PV;

            DataContext = this;
        }

        protected void OnPropertyChanged(string name) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));

        #region FILTER LOADERS
        private void LoadActiveParts()
        {
            ActiveParts.Clear();
            ActiveParts.Add("All");

            var parts = _dataService.GetActiveParts();
            if (parts != null)
                foreach (var p in parts)
                    if (!string.IsNullOrWhiteSpace(p.Para_No))
                        ActiveParts.Add(p.Para_No);

            SelectedPartNo = ActiveParts.FirstOrDefault() ?? "All";
        }

        private async Task OnSelectedPartNoChangedAsync()
        {
            await LoadLotNumbersAsync(SelectedPartNo);
            LoadParameters(SelectedPartNo);
            await LoadOperatorsAsync(SelectedPartNo);
        }

        public async Task LoadLotNumbersAsync(string partNo)
        {
            var lots = await _dataService.GetLotNumbersByPartAndDateRangeAsync(
                partNo == "All" ? null : partNo, SelectedDateTimeFrom, SelectedDateTimeTo);

            LotNumbers.Clear();
            LotNumbers.Add("All");

            if (lots != null)
                foreach (var lot in lots)
                    if (!string.IsNullOrWhiteSpace(lot))
                        LotNumbers.Add(lot);

            SelectedLotNo = LotNumbers.FirstOrDefault() ?? "All";
        }

        private void LoadParameters(string? part)
        {
            ParametersOptions.Clear();

            string partFilter = part == "All" ? string.Empty : part ?? string.Empty;
            var config = data_service_safe_getpartconfig(partFilter);

            if (config != null)
                foreach (var c in config) ParametersOptions.Add(c.Parameter);

            SelectedParameter = ParametersOptions.FirstOrDefault();
        }


        // small safe accessor in case _dataService is null (defensive)
        private IEnumerable<dynamic> data_service_safe_getpartconfig(string part) =>
            _dataService?.GetPartConfig(part) ?? Enumerable.Empty<dynamic>();

        private void LoadOperators()
        {
            Operators.Clear();
            Operators.Add("All");
            SelectedOperator = "All";
        }

        private async Task LoadOperatorsAsync(string part)
        {
            Operators.Clear();
            Operators.Add("All");

            var ops = await _dataService.GetOperatorsByPartAndDateRangeAsync(
                part == "All" ? null : part, SelectedDateTimeFrom, SelectedDateTimeTo);

            if (ops != null)
                foreach (var op in ops)
                    if (!string.IsNullOrWhiteSpace(op))
                        Operators.Add(op);

            SelectedOperator = Operators.FirstOrDefault() ?? "All";
        }



        private async Task ReloadLotAsync() => await LoadLotNumbersAsync(SelectedPartNo);
        #endregion

        private bool CanSubmit() =>
            SelectedDateTimeFrom != null &&
            SelectedDateTimeTo != null &&
            !string.IsNullOrWhiteSpace(SelectedParameter);

        private void ClosePage()
        {
            var window = Window.GetWindow(this);
            // optional: window?.Close();
        }

        #region MAIN METHOD
        private async Task LoadRnRDataAsync()
        {
            try
            {
                ClearAllGrids();
                ResetSummary();

                if (string.IsNullOrWhiteSpace(SelectedParameter))
                {
                    MessageBox.Show("Please select a parameter.");
                    return;
                }

                string? part = SelectedPartNo == "All" ? null : SelectedPartNo;
                string? lot = SelectedLotNo == "All" ? null : SelectedLotNo;
                string? oper = SelectedOperator == "All" ? null : SelectedOperator;

                var config = _dataService.GetPartConfig(part)
                                         ?.FirstOrDefault(c => c.Parameter == SelectedParameter);

                if (config == null)
                {
                    MessageBox.Show("Part configuration not found.");
                    return;
                }

                string? col = ParameterToColumn.ContainsKey(SelectedParameter)
                    ? ParameterToColumn[SelectedParameter]
                    : SelectedParameter;

                var list = await _dataService.GetMeasurementReadingsAsync(
                    part, lot, oper, SelectedDateTimeFrom, SelectedDateTimeTo);

                if (list == null || list.Count == 0)
                {
                    MessageBox.Show("No measurement data found.");
                    return;
                }

                // Extract numeric values safely (handles numeric types)
                List<double> values = new();
                foreach (var m in list)
                {
                    var prop = m.GetType().GetProperty(col);
                    if (prop == null) continue;
                    var raw = prop.GetValue(m);
                    if (raw == null) continue;

                    if (raw is double d) values.Add(d);
                    else if (raw is float f) values.Add(Convert.ToDouble(f));
                    else if (raw is decimal dec) values.Add(Convert.ToDouble(dec));
                    else if (raw is int i) values.Add(Convert.ToDouble(i));
                    else if (double.TryParse(raw.ToString(), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double parsed))
                        values.Add(parsed);
                }

                if (values.Count < 90)
                {
                    MessageBox.Show($"Need at least 90 measurements. Found: {values.Count}");
                    return;
                }

                values = values.Take(90).ToList();

                var a1 = values.Skip(0).Take(30).ToList();
                var a2 = values.Skip(30).Take(30).ToList();
                var a3 = values.Skip(60).Take(30).ToList();

                BuildGridForAppraiser(Appraiser1Data, a1);
                BuildGridForAppraiser(Appraiser2Data, a2);
                BuildGridForAppraiser(Appraiser3Data, a3);

                double dataMin = values.Min();
                double dataMax = values.Max();
                double toleranceFromData = dataMax - dataMin;   // your “actual” tolerance from readings

                TotalTolerance = toleranceFromData;
                //-----------------------------------------------------------
                // Compute both method results and store them
                //-----------------------------------------------------------
                var both = ComputeBothMethods(a1, a2, a3, TotalTolerance);

                PvMethodResult = both.Item1;
                ToleranceMethodResult = both.Item2;

                // By default show PV method (or keep prior selection)
                if (SelectedMethod == RrMethod.PV)
                    SelectedMethod = RrMethod.PV; // force update
                else
                    UpdateDisplayedResults();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        #endregion

        #region RESULT DTO
        // Single DTO used to hold method outputs
        public class GaugeRrResult
        {
            public double Ev { get; set; }                 // Equipment Variation
            public double Av { get; set; }                 // Appraiser Variation
            public double Grr { get; set; }                // Repeatability + Reproducibility
            public double Pv { get; set; }                 // Part Variation
            public double Tv { get; set; }                 // Total Variation (PV method) or Tol/6 (Tolerance method)
            public double PercentGrr_Pv { get; set; }      // %GRR relative to TV
            public double PercentPv { get; set; }          // %PV relative to TV
            public double Ndc { get; set; }                // Number of distinct categories
            public double PercentGrrTolerance { get; set; } // %GRR relative to Tol/6 (tolerance method)
            public string? ConclusionTolerance { get; set; } // Conclusion based on tolerance method
        }
        #endregion

        #region CALCULATIONS (both methods)

        private (GaugeRrResult pvResult, GaugeRrResult tolResult) ComputeBothMethods(
     List<double> a1, List<double> a2, List<double> a3, double tolerance)
        {
            // 1. Convert lists to 3x10 matrices
            var m1 = ConvertToMatrix(a1);
            var m2 = ConvertToMatrix(a2);
            var m3 = ConvertToMatrix(a3);

            // 2. Xbar and Rbar for each appraiser
            double X1Bar = 0, X2Bar = 0, X3Bar = 0;
            double R1Bar = 0, R2Bar = 0, R3Bar = 0;

            for (int p = 0; p < 10; p++)
            {
                double[] v1 = { m1[0, p], m1[1, p], m1[2, p] };
                double[] v2 = { m2[0, p], m2[1, p], m2[2, p] };
                double[] v3 = { m3[0, p], m3[1, p], m3[2, p] };

                X1Bar += v1.Average(); R1Bar += v1.Max() - v1.Min();
                X2Bar += v2.Average(); R2Bar += v2.Max() - v2.Min();
                X3Bar += v3.Average(); R3Bar += v3.Max() - v3.Min();
            }

            X1Bar /= 10; X2Bar /= 10; X3Bar /= 10;
            R1Bar /= 10; R2Bar /= 10; R3Bar /= 10;

            double RDoubleBar = (R1Bar + R2Bar + R3Bar) / 3.0;

            double R = Math.Round(RDoubleBar, 4);
            // 3. Constants (VB same)
            const double k1 = 0.5908;
            const double k2 = 0.5231;
            const double k3 = 0.31;
            const double ndcFactor = 1.41;

            // 4. EV
            double EV = Math.Round(R * k1, 4);

            // 5. AV
            double L29 = Math.Max(X1Bar, Math.Max(X2Bar, X3Bar))
                       - Math.Min(X1Bar, Math.Min(X2Bar, X3Bar));

            double B54 = (EV * EV) / 30.0;
            double B53 = Math.Pow(L29 * k2, 2);

            double AV = Math.Round(Math.Sqrt(Math.Abs(B53 - B54)), 4);

            // 6. GRR
            double GRR = Math.Round(Math.Sqrt(EV * EV + AV * AV), 6);

            // 7. PV
            double[] partAvgs = new double[10];
            for (int p = 0; p < 10; p++)
            {
                double[] allVals =
                {
            m1[0,p], m1[1,p], m1[2,p],
            m2[0,p], m2[1,p], m2[2,p],
            m3[0,p], m3[1,p], m3[2,p]
        };
                partAvgs[p] = allVals.Average();
            }

            Array.Sort(partAvgs);
            double Rp = partAvgs[9] - partAvgs[0];

            double PV = Math.Round(Rp * k3, 4);

            // 8. PV Method
            double TV_PV = Math.Round(Math.Sqrt(GRR * GRR + PV * PV), 7);

            double percentGrr_PV = Math.Round((GRR / TV_PV) * 100.0, 4);
            double percentPV = Math.Round((PV / TV_PV) * 100.0, 4);

            double NDC = Math.Round((PV / GRR) * ndcFactor, 2);

            //string conclusionPV =
            //    percentGrr_PV <= 10 ? "Accepted" :
            //    percentGrr_PV <= 30 ? "Conditionally Accepted" :
            //    "Not Accepted";

            var pvResult = new GaugeRrResult
            {
                Ev = EV,
                Av = AV,
                Grr = GRR,
                Pv = PV,
                Tv = TotalTolerance,
                PercentGrr_Pv = percentGrr_PV,
                PercentPv = percentPV,
                Ndc = NDC,
                PercentGrrTolerance = 0,
                //ConclusionTolerance = conclusionPV
            };

            // 9. Tolerance Method
            double TV_Tol = Math.Round(TotalTolerance, 4);
            double percentGrr_Tol = Math.Round((GRR / TV_Tol) * 100.0, 4);
            double percentPV_Tol = Math.Round((PV / TV_Tol) * 100.0, 4);

            string conclusionTol =
                percentGrr_Tol <= 10 ? "Accepted" :
                percentGrr_Tol <= 30 ? "Conditionally Accepted" :
                "Not Accepted";

            var tolResult = new GaugeRrResult
            {
                Ev = EV,
                Av = AV,
                Grr = GRR,
                Pv = PV,
                Tv = TotalTolerance,
                PercentGrr_Pv = percentPV_Tol,
                PercentPv = percentPV_Tol,
                Ndc = NDC,
                PercentGrrTolerance = percentGrr_Tol,
                ConclusionTolerance = conclusionTol
            };

            return (pvResult, tolResult);
        }


        #endregion

        #region HELPERS
        private double[,] ConvertToMatrix(List<double> values)
        {
            var m = new double[3, 10];
            for (int t = 0; t < 3; t++)
                for (int p = 0; p < 10; p++)
                    m[t, p] = values[t * 10 + p];
            return m;
        }

        private void BuildGridForAppraiser(ObservableCollection<RnRGridRow> target, List<double> vals)
        {
            target.Clear();

            for (int t = 0; t < 3; t++)
            {
                var row = new RnRGridRow
                {
                    TrialNo = (t + 1).ToString(),
                    Data1 = vals[t * 10 + 0],
                    Data2 = vals[t * 10 + 1],
                    Data3 = vals[t * 10 + 2],
                    Data4 = vals[t * 10 + 3],
                    Data5 = vals[t * 10 + 4],
                    Data6 = vals[t * 10 + 5],
                    Data7 = vals[t * 10 + 6],
                    Data8 = vals[t * 10 + 7],
                    Data9 = vals[t * 10 + 8],
                    Data10 = vals[t * 10 + 9],
                };
                target.Add(row);
            }

            var avg = new RnRGridRow { TrialNo = "Average" };
            var rng = new RnRGridRow { TrialNo = "Range" };

            for (int c = 0; c < 10; c++)
            {
                double[] column =
                {
                    target[0].GetColumn(c),
                    target[1].GetColumn(c),
                    target[2].GetColumn(c)
                };

                avg.SetColumn(c, column.Average());
                rng.SetColumn(c, column.Max() - column.Min());
            }

            target.Add(avg);
            target.Add(rng);
        }

        private void ClearAllGrids()
        {
            Appraiser1Data.Clear();
            Appraiser2Data.Clear();
            Appraiser3Data.Clear();
        }

        private void ResetSummary()
        {
            Repeatability = 0;
            Reproducibility = 0;
            PartVariation = 0;
            TotalPVPercent = 0;
            TotalTolerance = 0;
            GRRPercentage = 0;
            GRRConclusion = string.Empty;
            Ndc = 0;
            PvMethodResult = null;
            ToleranceMethodResult = null;
        }

        public class RelayCommand : ICommand
        {
            private readonly Func<object, Task> _async;
            private readonly Action<object> _sync;
            private readonly Predicate<object> _canExecute;

            public RelayCommand(Func<object, Task> async, Predicate<object> can = null) { _async = async; _canExecute = can; }
            public RelayCommand(Action<object> sync, Predicate<object> can = null) { _sync = sync; _canExecute = can; }

            public bool CanExecute(object parameter) => _canExecute == null || _canExecute(parameter);

            public async void Execute(object parameter)
            {
                if (_async != null)
                {
                    try { await _async(parameter); }
                    catch (Exception ex) { MessageBox.Show(ex.Message, "Command error", MessageBoxButton.OK, MessageBoxImage.Error); }
                }
                else _sync?.Invoke(parameter);
            }

            public event EventHandler CanExecuteChanged
            {
                add { CommandManager.RequerySuggested += value; }
                remove { CommandManager.RequerySuggested -= value; }
            }
        }

        public class RnRGridRow : INotifyPropertyChanged
        {
            public event PropertyChangedEventHandler PropertyChanged;


            public bool IsRangeRow => TrialNo == "Range";


            private string _trialNo;
            public string TrialNo { get => _trialNo; set { _trialNo = value; PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(TrialNo))); } }

            private double _d1, _d2, _d3, _d4, _d5, _d6, _d7, _d8, _d9, _d10;
            public double Data1 { get => _d1; set { _d1 = value; PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Data1))); } }
            public double Data2 { get => _d2; set { _d2 = value; PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Data2))); } }
            public double Data3 { get => _d3; set { _d3 = value; PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Data3))); } }
            public double Data4 { get => _d4; set { _d4 = value; PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Data4))); } }
            public double Data5 { get => _d5; set { _d5 = value; PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Data5))); } }
            public double Data6 { get => _d6; set { _d6 = value; PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Data6))); } }
            public double Data7 { get => _d7; set { _d7 = value; PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Data7))); } }
            public double Data8 { get => _d8; set { _d8 = value; PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Data8))); } }
            public double Data9 { get => _d9; set { _d9 = value; PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Data9))); } }
            public double Data10 { get => _d10; set { _d10 = value; PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Data10))); } }

            public double GetColumn(int c) =>
                c switch
                {
                    0 => Data1,
                    1 => Data2,
                    2 => Data3,
                    3 => Data4,
                    4 => Data5,
                    5 => Data6,
                    6 => Data7,
                    7 => Data8,
                    8 => Data9,
                    9 => Data10,
                    _ => 0
                };

            public double MaxRangeValue =>
        new[] { Data1, Data2, Data3, Data4, Data5,
                Data6, Data7, Data8, Data9, Data10 }.Max();

            public bool IsRangeAboveLimit =>
        IsRangeRow && MaxRangeValue > 0.010;
            public void SetColumn(int c, double v)
            {
                switch (c)
                {
                    case 0: Data1 = v; break;
                    case 1: Data2 = v; break;
                    case 2: Data3 = v; break;
                    case 3: Data4 = v; break;
                    case 4: Data5 = v; break;
                    case 5: Data6 = v; break;
                    case 6: Data7 = v; break;
                    case 7: Data8 = v; break;
                    case 8: Data9 = v; break;
                    case 9: Data10 = v; break;
                }
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs($"Data{c + 1}"));
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(MaxRangeValue)));
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsRangeAboveLimit)));

            }
        }
        #endregion


        private List<double> GetRowValues(RnRGridRow row) =>
    new()
    {
        row.Data1, row.Data2, row.Data3, row.Data4, row.Data5,
        row.Data6, row.Data7, row.Data8, row.Data9, row.Data10
    };

        private void RecalculateFromGrids()
        {
            // 1) rebuild averages and ranges from trials
            RebuildAvgAndRangeRows(Appraiser1Data);
            RebuildAvgAndRangeRows(Appraiser2Data);
            RebuildAvgAndRangeRows(Appraiser3Data);

            // 2) read 3 trial rows back into vectors
            var a1 = Appraiser1Data.Take(3).SelectMany(GetRowValues).ToList();
            var a2 = Appraiser2Data.Take(3).SelectMany(GetRowValues).ToList();
            var a3 = Appraiser3Data.Take(3).SelectMany(GetRowValues).ToList();

            if (a1.Count != 30 || a2.Count != 30 || a3.Count != 30)
                return;

            var all = a1.Concat(a2).Concat(a3).ToList();
            double dataMin = all.Min();
            double dataMax = all.Max();
            TotalTolerance = Math.Round(dataMax - dataMin, 4);

            var both = ComputeBothMethods(a1, a2, a3, TotalTolerance);
            PvMethodResult = both.pvResult;
            ToleranceMethodResult = both.tolResult;
            UpdateDisplayedResults();
        }

        private void RebuildAvgAndRangeRows(ObservableCollection<RnRGridRow> target)
        {
            if (target.Count < 5) return;

            var r1 = target[0];
            var r2 = target[1];
            var r3 = target[2];
            var avg = target[3];
            var rng = target[4];

            for (int c = 0; c < 10; c++)
            {
                double[] col =
                {
            r1.GetColumn(c),
            r2.GetColumn(c),
            r3.GetColumn(c)
        };

                avg.SetColumn(c, col.Average());
                rng.SetColumn(c, col.Max() - col.Min());
            }
        }

        private string GetWritableLogoFolder()
        {
            string baseFolder = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            string logoFolder = System.IO.Path.Combine(baseFolder, "EVMS", "CompanyLogo");

            if (!Directory.Exists(logoFolder))
                Directory.CreateDirectory(logoFolder);

            return logoFolder;
        }


        // Global variables for company info
        private string _companyName = "";
        private string _companyLogoPath = "";

        private void LoadCompanyInfo()
        {
            string query = "SELECT TOP 1 CompanyName, LogoPath FROM CompanyConfig WHERE Id = 1";

            using SqlConnection conn = new SqlConnection(connectionString);
            using SqlCommand cmd = new SqlCommand(query, conn);
            conn.Open();

            using SqlDataReader reader = cmd.ExecuteReader();
            if (reader.Read())
            {
                _companyName = reader["CompanyName"]?.ToString() ?? "";
                //txtCompanyName.Text = _companyName;

                string logoFile = reader["LogoPath"]?.ToString();

                if (!string.IsNullOrEmpty(logoFile))
                {
                    // FIX: use a writable local folder, NOT app directory
                    string logoFolder = GetWritableLogoFolder();
                    _companyLogoPath = System.IO.Path.Combine(logoFolder, logoFile);

                    if (File.Exists(_companyLogoPath))
                    {
                        //imgPreview.Source = new BitmapImage(new Uri(_companyLogoPath));
                    }
                }
            }
        }
        private void ExportToExcel()
        {
            try
            {
                if (!Appraiser1Data.Any() && !Appraiser2Data.Any() && !Appraiser3Data.Any())
                {
                    MessageBox.Show("No data to export. Please generate R&R first.");
                    return;
                }

                string folder = @"D:\MEPL\RnR Report";
                if (!Directory.Exists(folder))
                    Directory.CreateDirectory(folder);

                string fileName = $"RnR_{SelectedPartNo}_{SelectedParameter}_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
                string filePath = Path.Combine(folder, fileName);

                using (var wb = new XLWorkbook())
                {
                    var ws = wb.Worksheets.Add("R&R Report");

                    // ===============================================================
                    // HEADER – LOGO + COMPANY NAME
                    // ===============================================================
                    ws.Row(1).Height = 50;
                    ws.Row(2).Height = 25;
                    ws.Row(3).Height = 15;

                    ws.Range("A1:A2").Merge();
                    if (!string.IsNullOrEmpty(_companyLogoPath) && File.Exists(_companyLogoPath))
                    {
                        try
                        {
                            ws.AddPicture(_companyLogoPath)
                              .MoveTo(ws.Cell("A1"), 3, 3)
                              .Scale(0.35);
                        }
                        catch { }
                    }
                    ws.Column(1).Width = 12;

                    ws.Range("B1:G1").Merge();
                    var companyCell = ws.Cell("B1");
                    companyCell.Value = _companyName ?? "COMPANY NAME";
                    companyCell.Style.Font.SetBold();
                    companyCell.Style.Font.SetFontSize(14);
                    companyCell.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
                    companyCell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Bottom;

                    // ===============================================================
                    // REPORT TITLE
                    // ===============================================================
                    ws.Row(4).Height = 25;
                    ws.Row(5).Height = 25;

                    ws.Range("A3:P4").Merge();
                    var titleCell = ws.Cell("A3");
                    titleCell.Value = "R&R REPORT";
                    titleCell.Style.Font.SetBold();
                    titleCell.Style.Font.SetFontSize(14);
                    titleCell.Style.Font.SetFontColor(XLColor.White);
                    titleCell.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                    titleCell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    titleCell.Style.Fill.BackgroundColor = XLColor.FromHtml("#1F4E78");

                    // ===============================================================
                    // INFO SECTION
                    // ===============================================================
                    ws.Row(6).Height = 18;
                    ws.Row(7).Height = 18;
                    ws.Row(8).Height = 18;

                    ws.Cell("A6").Value = "Part No:"; ws.Cell("A6").Style.Font.SetBold();
                    ws.Cell("C6").Value = SelectedPartNo;

                    ws.Cell("E6").Value = "Lot No:"; ws.Cell("E6").Style.Font.SetBold();
                    ws.Cell("G6").Value = SelectedLotNo;

                    ws.Cell("J6").Value = "Parameter:"; ws.Cell("J6").Style.Font.SetBold();
                    ws.Cell("L6").Value = SelectedParameter;

                    ws.Cell("A7").Value = "Operator:"; ws.Cell("A7").Style.Font.SetBold();
                    ws.Cell("C7").Value = SelectedOperator;

                    ws.Cell("E7").Value = "From:"; ws.Cell("E7").Style.Font.SetBold();
                    ws.Cell("G7").Value = SelectedDateTimeFrom?.ToString("dd-MMM-yyyy");

                    ws.Cell("J7").Value = "To:"; ws.Cell("J7").Style.Font.SetBold();
                    ws.Cell("L7").Value = SelectedDateTimeTo?.ToString("dd-MMM-yyyy");

                    ws.Range("A8:P8").Style.Border.BottomBorder = XLBorderStyleValues.Medium;

                    int startRow = 10;

                    // ===============================================================
                    // LEFT SIDE – APPRAISERS
                    // ===============================================================
                    int leftRow = startRow;
                    leftRow = WriteAppraiserReadings(ws, leftRow, "APPRAISER 1", Appraiser1Data);
                    leftRow = WriteAppraiserReadings(ws, leftRow, "APPRAISER 2", Appraiser2Data);
                    leftRow = WriteAppraiserReadings(ws, leftRow, "APPRAISER 3", Appraiser3Data);

                    // ===============================================================
                    // RIGHT SIDE – CALCULATION METHODS
                    int calcRow = startRow;
                    int calcCol = 13; // Column M

                    // PV METHOD
                    ws.Range(calcRow, calcCol, calcRow, calcCol + 3).Merge();
                    ws.Cell(calcRow, calcCol).Value = "PV METHOD RESULTS";
                    ws.Cell(calcRow, calcCol).Style.Font.SetBold().Font.SetFontColor(XLColor.White);
                    ws.Cell(calcRow, calcCol).Style.Fill.BackgroundColor = XLColor.SteelBlue;
                    ws.Cell(calcRow, calcCol).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                    calcRow++;

                    if (PvMethodResult != null)
                    {
                        calcRow = WriteResult(ws, calcRow, "Repeatability (EV)", PvMethodResult.Ev, calcCol);
                        calcRow = WriteResult(ws, calcRow, "Reproducibility (AV)", PvMethodResult.Av, calcCol);
                        calcRow = WriteResult(ws, calcRow, "GRR", PvMethodResult.Grr, calcCol);
                        calcRow = WriteResult(ws, calcRow, "Part Variation (PV)", PvMethodResult.Pv, calcCol);
                        calcRow = WriteResult(ws, calcRow, "Total Variation (TV)", PvMethodResult.Tv, calcCol);
                        calcRow = WriteResult(ws, calcRow, "%GRR", PvMethodResult.PercentGrr_Pv, calcCol);
                        calcRow = WriteResult(ws, calcRow, "NDC", PvMethodResult.Ndc, calcCol);
                    }

                    calcRow++;

                    // TOLERANCE METHOD
                    ws.Range(calcRow, calcCol, calcRow, calcCol + 3).Merge();
                    ws.Cell(calcRow, calcCol).Value = "TOLERANCE METHOD RESULTS";
                    ws.Cell(calcRow, calcCol).Style.Font.SetBold().Font.SetFontColor(XLColor.White);
                    ws.Cell(calcRow, calcCol).Style.Fill.BackgroundColor = XLColor.SeaGreen;
                    ws.Cell(calcRow, calcCol).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                    calcRow++;

                    if (ToleranceMethodResult != null)
                    {
                        calcRow = WriteResult(ws, calcRow, "Repeatability (EV)", ToleranceMethodResult.Ev, calcCol);
                        calcRow = WriteResult(ws, calcRow, "Reproducibility (AV)", ToleranceMethodResult.Av, calcCol);
                        calcRow = WriteResult(ws, calcRow, "GRR", ToleranceMethodResult.Grr, calcCol);
                        calcRow = WriteResult(ws, calcRow, "Total Tolerance", ToleranceMethodResult.Tv, calcCol);
                        calcRow = WriteResult(ws, calcRow, "%GRR", ToleranceMethodResult.PercentGrrTolerance, calcCol);
                        calcRow = WriteResult(ws, calcRow, "NDC", ToleranceMethodResult.Ndc, calcCol);
                        calcRow = WriteResult(ws, calcRow, "Conclusion", ToleranceMethodResult.ConclusionTolerance, calcCol);
                    }


                    ws.Columns().AdjustToContents();
                    wb.SaveAs(filePath);
                }

                Process.Start(new ProcessStartInfo(filePath) { UseShellExecute = true });
                MessageBox.Show("Excel R&R report generated successfully.");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Export Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }




        private int WriteAppraiserReadings(
     IXLWorksheet ws,
     int startRow,
     string title,
     ObservableCollection<RnRGridRow> data)
        {
            // ===== APPRAISER TITLE =====
            ws.Cell(startRow, 1).Value = title;
            ws.Range(startRow, 1, startRow, 11).Merge()
                .Style.Font.SetBold()
                .Fill.SetBackgroundColor(XLColor.RoyalBlue)
                .Font.SetFontColor(XLColor.White);

            startRow++;

            // ===== HEADER ROW =====
            ws.Cell(startRow, 1).Value = "Trial";
            for (int i = 1; i <= 10; i++)
                ws.Cell(startRow, i + 1).Value = i;

            ws.Range(startRow, 1, startRow, 11).Style
                .Font.SetBold()
                .Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);

            startRow++;

            // ===== WRITE ALL ROWS (Trial 1,2,3, Average, Range) =====
            foreach (var r in data)
            {
                ws.Cell(startRow, 1).Value = r.TrialNo;
                ws.Cell(startRow, 1).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);

                double[] values =
                {
            r.Data1, r.Data2, r.Data3, r.Data4, r.Data5,
            r.Data6, r.Data7, r.Data8, r.Data9, r.Data10
        };

                for (int i = 0; i < 10; i++)
                {
                    ws.Cell(startRow, i + 2).Value = values[i];
                    ws.Cell(startRow, i + 2).Style.NumberFormat.Format = "0.0000";
                    ws.Cell(startRow, i + 2).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                }

                // Shade Average & Range rows
                if (r.TrialNo == "Average" || r.TrialNo == "Range")
                {
                    ws.Range(startRow, 1, startRow, 11)
                      .Style.Fill.SetBackgroundColor(XLColor.LightGray);
                    ws.Cell(startRow, 1).Style.Font.SetBold();
                }

                startRow++;
            }

            // ===== BORDERS =====
            ws.Range(startRow - data.Count - 1, 1, startRow - 1, 11)
                .Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

            ws.Range(startRow - data.Count - 1, 1, startRow - 1, 11)
                .Style.Border.InsideBorder = XLBorderStyleValues.Thin;

            startRow++; // space after table
            return startRow;
        }




        private int WriteInfo(IXLWorksheet ws, int row, string label, object value)
        {
            ws.Cell(row, 1).Value = label;
            ws.Cell(row, 1).Style.Font.SetBold();

            if (value == null)
                ws.Cell(row, 2).Value = "";
            else if (value is DateTime dt)
                ws.Cell(row, 2).Value = dt;
            else if (value is double d)
                ws.Cell(row, 2).Value = d;
            else if (value is int i)
                ws.Cell(row, 2).Value = i;
            else
                ws.Cell(row, 2).Value = value.ToString();

            return row + 1;
        }


        private int WriteResult(IXLWorksheet ws, int row, string label, object value, int startCol)
        {
            ws.Cell(row, startCol).Value = label;
            ws.Cell(row, startCol).Style.Font.SetBold();

            if (value == null)
                ws.Cell(row, startCol + 3).Value = "";
            else if (value is double d)
                ws.Cell(row, startCol + 3).Value = d;
            else if (value is int i)
                ws.Cell(row, startCol + 3).Value = i;
            else
                ws.Cell(row, startCol + 3).Value = value.ToString();

            ws.Cell(row, startCol + 3).Style.NumberFormat.Format = "0.0000";
            return row + 1;
        }


    }
}
