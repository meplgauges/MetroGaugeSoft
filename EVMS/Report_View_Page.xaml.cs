using ClosedXML.Excel;
using EVMS.Service;
using Microsoft.Data.SqlClient;
using PdfSharpCore.Drawing;
using PdfSharpCore.Pdf;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using static EVMS.Login_Page;

namespace EVMS
{
    public partial class Report_View_Page : UserControl, INotifyPropertyChanged
    {
        private readonly DataStorageService _dataService;
        private readonly string connectionString;

        public ObservableCollection<string> ActiveParts { get; set; } = new();
        public ObservableCollection<string> LotNumbers { get; set; } = new();
        public ObservableCollection<string> ParametersOptions { get; set; } = new();
        public ObservableCollection<string> DesignOptions { get; set; } = new();
        public ObservableCollection<string> Operators { get; set; } = new();
        public ObservableCollection<ReportTableItem> ReportTableItems { get; set; } = new();

        public ObservableCollection<string> NgOkCountOptions { get; set; } = new ObservableCollection<string> { "No", "Yes" };

        public ObservableCollection<string> ReportTypeOptions { get; set; } = new ObservableCollection<string> { "All", "NG", "OK" };

        public ObservableCollection<NgOkSummaryItem> NgOkSummaryItems { get; set; } = new ObservableCollection<NgOkSummaryItem>();


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

                    // Reload parameters whenever selected part changes
                    LoadParameters();

                    // Also reload lot and operator asynchronously
                    _ = ReloadLotAndOperatorAsync();
                }
            }
        }


        private string _selectedLotNo;
        public string SelectedLotNo
        {
            get => _selectedLotNo;
            set { _selectedLotNo = value; OnPropertyChanged(nameof(SelectedLotNo)); }
        }

        private string _selectedParameter;
        public string SelectedParameter
        {
            get => _selectedParameter;
            set { _selectedParameter = value; OnPropertyChanged(nameof(SelectedParameter)); }
        }


        private string _selectedOperator;
        public string SelectedOperator
        {
            get => _selectedOperator;
            set { _selectedOperator = value; OnPropertyChanged(nameof(SelectedOperator)); }
        }

        private DateTime? _selectedDateTimeFrom = DateTime.Now.AddDays(-7);
        public DateTime? SelectedDateTimeFrom
        {
            get => _selectedDateTimeFrom;
            set
            {
                if (_selectedDateTimeFrom != value)
                {
                    _selectedDateTimeFrom = value;
                    OnPropertyChanged(nameof(SelectedDateTimeFrom));
                    _ = ReloadLotAndOperatorAsync();
                }
            }
        }

        private string _showNgOkSummary = "No";
        public string ShowNgOkSummary
        {
            get => _showNgOkSummary;
            set
            {
                if (_showNgOkSummary != value)
                {
                    _showNgOkSummary = value;
                    OnPropertyChanged(nameof(ShowNgOkSummary));
                }
            }
        }


        private string selectedReportType = "All";
        public string SelectedReportType
        {
            get => selectedReportType;
            set
            {
                if (selectedReportType != value)
                {
                    selectedReportType = value;
                    OnPropertyChanged(nameof(SelectedReportType));
                }
            }
        }

        private DateTime? _selectedDateTimeTo = DateTime.Now;
        public DateTime? SelectedDateTimeTo
        {
            get => _selectedDateTimeTo;
            set
            {
                if (_selectedDateTimeTo != value)
                {
                    _selectedDateTimeTo = value;
                    OnPropertyChanged(nameof(SelectedDateTimeTo));
                    _ = ReloadLotAndOperatorAsync();
                }
            }
        }

        public Report_View_Page()
        {
            InitializeComponent();
            _dataService = new DataStorageService();
            connectionString = ConfigurationManager.ConnectionStrings["EVMSDb"].ConnectionString;

            LoadDesignOptions();
            LoadActiveParts();
            LoadParameters();
            LoadCompanyInfo();
            this.Loaded += ReportViewPage_Loaded;
            this.PreviewKeyDown += ReportViewPage_PreviewKeyDown;

            DataContext = this;
        }

        private void ReportViewPage_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                HandleEscKeyAction();
                e.Handled = true;
            }
        }

        private void ReportViewPage_Loaded(object? sender, RoutedEventArgs e)
        {
            this.Focusable = true;
            this.IsTabStop = true;
            Keyboard.Focus(this);
            FocusManager.SetFocusedElement(Window.GetWindow(this)!, this);
        }

        private void HandleEscKeyAction()
        {
            Window currentWindow = Window.GetWindow(this);
            if (currentWindow != null)
            {
                var mainContentGrid = currentWindow.FindName("MainContentGrid") as Grid;
                if (mainContentGrid != null)
                {
                    mainContentGrid.Children.Clear();

                    var resultPage = new Dashboard
                    {
                        HorizontalAlignment = HorizontalAlignment.Stretch,
                        VerticalAlignment = VerticalAlignment.Stretch
                    };

                    mainContentGrid.Children.Add(resultPage);
                }
            }
        }



        private void LoadDesignOptions()
        {
            DesignOptions.Clear();
            DesignOptions.Add("Table View");
        }

        private void LoadActiveParts()
        {
            var parts = _dataService?.GetActiveParts();
            ActiveParts.Clear();
            ActiveParts.Add("All");

            foreach (var part in parts)
            {
                if (!string.IsNullOrWhiteSpace(part.Para_No))
                    ActiveParts.Add(part.Para_No);
            }

            SelectedPartNo = ActiveParts.FirstOrDefault();
        }

        private void LoadParameters()
        {
            ParametersOptions.Clear();
            ParametersOptions.Add("All");

            if (string.IsNullOrEmpty(SelectedPartNo)) return;

            var exampleParams = _dataService
                .GetPartConfig(SelectedPartNo)
                .Select(p => p.Parameter)
                .Distinct()
                .Where(p => !string.IsNullOrWhiteSpace(p))
                .ToList();

            foreach (var p in exampleParams)
                ParametersOptions.Add(p);

            SelectedParameter = ParametersOptions.FirstOrDefault() ?? "All";

            // Raise PropertyChanged for ParametersOptions and SelectedParameter if using MVVM
            OnPropertyChanged(nameof(ParametersOptions));
            OnPropertyChanged(nameof(SelectedParameter));
        }


        /// <summary>
        /// Updates Lot Numbers and Operators based on selected Part and Date Range.
        /// </summary>
        private async Task ReloadLotAndOperatorAsync()
        {
            try
            {
                string? partFilter = SelectedPartNo;
                if (partFilter == "All")
                {
                    partFilter = null;  // or empty string, depending on your data layer
                }
                DateTime? from = SelectedDateTimeFrom;
                DateTime? to = SelectedDateTimeTo;

                LotNumbers.Clear();
                LotNumbers.Add("All");
                var lots = await _dataService.GetLotNumbersByPartAndDateRangeAsync(partFilter, from, to);
                foreach (var lot in lots)
                    LotNumbers.Add(lot);
                SelectedLotNo = LotNumbers.FirstOrDefault();

                Operators.Clear();
                Operators.Add("All");
                var ops = await _dataService.GetOperatorsByPartAndDateRangeAsync(partFilter, from, to);
                foreach (var op in ops)
                    Operators.Add(op);
                SelectedOperator = Operators.FirstOrDefault();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading Lot/Operator data: {ex.Message}");
            }
        }

        private async void OnSubmitClicked(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(SelectedPartNo))
            {
                MessageBox.Show("Please select a Part No. If you want all parts, select 'All'.", "Input Required", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            await LoadDataBasedOnSelectionAsync();
        }


        private async Task LoadDataBasedOnSelectionAsync()
        {
            if (ShowNgOkSummary == "Yes")
            {
                await LoadNgOkCountTableAsync();
                ReportDataGrid.ItemsSource = NgOkSummaryItems;
            }
            else
            {
                await LoadReportTableAsync();
                ReportDataGrid.ItemsSource = ReportTableItems;
            }
        }



        private async Task LoadReportTableAsync()
        {
            ReportTableItems.Clear();

            string? partFilter = SelectedPartNo == "All" ? null : SelectedPartNo;
            string? lotFilter = SelectedLotNo == "All" ? null : SelectedLotNo;
            string? operatorFilter = SelectedOperator == "All" ? null : SelectedOperator;

            var results = await _dataService.GetMeasurementReadingsAsync(
                        partFilter ?? string.Empty,
                        lotFilter ?? string.Empty,
                        operatorFilter ?? string.Empty,
                        _selectedDateTimeFrom,
                        _selectedDateTimeTo);


            if (results != null && results.Any())
            {
                List<MeasurementReading> filteredResults = results.ToList();

                if (SelectedReportType == "NG")
                {
                    filteredResults = results.Where(r => r.Status == "NG").ToList();
                }
                else if (SelectedReportType == "OK")
                {
                    filteredResults = results.Where(r => r.Status == "OK").ToList();
                }

                if (!filteredResults.Any())
                {
                    MessageBox.Show("No data found for the selected filters.");
                    return;
                }

                var partConfigs = _dataService.GetPartConfig(partFilter ?? filteredResults.First().PartNo);
                var allParameters = partConfigs.Select(c => c.Parameter).Distinct().ToList();

                List<string> finalParams = SelectedParameter == "All"
                    ? allParameters
                    : new List<string> { SelectedParameter };

                GenerateParameterColumns(finalParams);

                int serialNo = 1;

                foreach (var r in filteredResults)
                {
                    var item = new ReportTableItem
                    {
                        Id = r.Id,
                        SerialNo = serialNo++,
                        PartNo = r.PartNo,
                        LotNo = r.LotNo,
                        Operator = r.Operator_ID,
                        Date = r.MeasurementDate.ToString("yyyy-MM-dd HH:mm"),
                        Parameters = new Dictionary<string, double>()
                    };

                    bool isAnyParamOutOfRange = false;

                    foreach (var param in finalParams)
                    {
                        var propName = MapParameterToProperty(param);
                        var propInfo = r.GetType().GetProperty(propName);

                        double measuredValue = 0;

                        if (propInfo != null)
                        {
                            var valueObj = propInfo.GetValue(r);
                            if (valueObj != null && double.TryParse(valueObj.ToString(), out double parsed))
                                measuredValue = parsed;
                        }

                        var config = partConfigs.FirstOrDefault(pc =>
                            pc.Parameter.Equals(param, StringComparison.OrdinalIgnoreCase));

                        if (config != null)
                        {
                            double usl = config.Nominal + config.RTolPlus;
                            double lsl = config.Nominal - config.RTolMinus;

                            if (measuredValue < lsl || measuredValue > usl)
                                isAnyParamOutOfRange = true;
                        }

                        item.Parameters[param] = measuredValue;
                    }

                    item.Status = isAnyParamOutOfRange ? "FAIL" : "PASS";

                    ReportTableItems.Add(item);
                }
            }
            else
            {
                MessageBox.Show("No data found for the selected filters.");
            }
        }


        private async Task LoadNgOkCountTableAsync()
        {
            NgOkSummaryItems.Clear();

            string? partFilter = SelectedPartNo == "All" ? null : SelectedPartNo;
            string? lotFilter = SelectedLotNo == "All" ? null : SelectedLotNo;
            string? operatorFilter = SelectedOperator == "All" ? null : SelectedOperator;

            var results = await _dataService.GetMeasurementReadingsAsync(
                partFilter, lotFilter, operatorFilter, _selectedDateTimeFrom, _selectedDateTimeTo);

            if (results == null || !results.Any())
            {
                MessageBox.Show("No data found for the selected filters.");
                return;
            }

            var partConfigs = _dataService.GetPartConfig(partFilter ?? results.First().PartNo);

            List<string> parametersToSummarize = SelectedParameter == "All"
                ? partConfigs.Select(pc => pc.Parameter).Distinct().ToList()
                : new List<string> { SelectedParameter };

            // Generate grid columns dynamically
            GenerateNgOkCountColumns(parametersToSummarize);

            int serialNo = 1;

            foreach (var param in parametersToSummarize)
            {
                var selectedParamConfig = partConfigs.FirstOrDefault(pc =>
                    pc.Parameter.Equals(param, StringComparison.OrdinalIgnoreCase));
                if (selectedParamConfig == null)
                    continue;

                var propName = MapParameterToProperty(param);

                double usl = selectedParamConfig.Nominal + selectedParamConfig.RTolPlus;
                double lsl = selectedParamConfig.Nominal - selectedParamConfig.RTolMinus;

                int ngCount = 0;
                int okCount = 0;

                foreach (var reading in results)
                {
                    var prop = reading.GetType().GetProperty(propName);
                    if (prop == null)
                        continue;

                    var valueObj = prop.GetValue(reading);
                    if (valueObj == null)
                        continue;

                    if (double.TryParse(valueObj.ToString(), out double val))
                    {
                        if (val < lsl || val > usl)
                            ngCount++;
                        else
                            okCount++;
                    }
                }

                var firstReading = results.First();

                NgOkSummaryItems.Add(new NgOkSummaryItem
                {
                    Id = firstReading.Id,
                    SerialNo = serialNo++,
                    PartNo = firstReading.PartNo,
                    LotNo = firstReading.LotNo,
                    Operator = firstReading.Operator_ID,
                    Date = firstReading.MeasurementDate.ToString("yyyy-MM-dd HH:mm"),
                    Parameter = param,
                    NgCount = ngCount,
                    OkCount = okCount,
                    Parameters = new Dictionary<string, double> { { param, ngCount + okCount } }
                });
            }
        }






        private void GenerateParameterColumns(List<string> parameters)
        {
            if (ReportDataGrid == null) return;

            // common centered style
            var centerStyle = new Style(typeof(TextBlock));
            centerStyle.Setters.Add(new Setter(TextBlock.TextAlignmentProperty, TextAlignment.Center));
            centerStyle.Setters.Add(new Setter(TextBlock.VerticalAlignmentProperty, VerticalAlignment.Center));

            // Keep fixed columns and remove existing parameter columns
            while (ReportDataGrid.Columns.Count > 5)
                ReportDataGrid.Columns.RemoveAt(5);

            foreach (var param in parameters)
            {
                var column = new DataGridTextColumn
                {
                    Header = param,
                    Binding = new Binding($"Parameters[{param}]"),
                    Width = 81,
                    ElementStyle = centerStyle
                };
                ReportDataGrid.Columns.Add(column);
            }

            // Add Status column always at the end
            ReportDataGrid.Columns.Add(new DataGridTextColumn
            {
                Header = "Status",
                Binding = new Binding("Status"),
                Width = 65,
                ElementStyle = centerStyle
            });
        }


        private void GenerateNgOkCountColumns(List<string> parameters)
        {
            if (ReportDataGrid == null) return;

            while (ReportDataGrid.Columns.Count > 5)
                ReportDataGrid.Columns.RemoveAt(5);

            ReportDataGrid.HorizontalContentAlignment = HorizontalAlignment.Stretch;

            DataGridTextColumn parameterColumn = new DataGridTextColumn
            {
                Header = "Parameter",
                Binding = new Binding("Parameter"),
                Width = new DataGridLength(2, DataGridLengthUnitType.Star) // occupies proportionally more space
            };
            ReportDataGrid.Columns.Add(parameterColumn);

            DataGridTextColumn ngColumn = new DataGridTextColumn
            {
                Header = "NG Count",
                Binding = new Binding("NgCount"),
                Width = new DataGridLength(1, DataGridLengthUnitType.Star),
                ElementStyle = new Style(typeof(TextBlock))
                {
                    Setters =
            {
                new Setter(TextBlock.ForegroundProperty, Brushes.Red),
                new Setter(TextBlock.FontWeightProperty, FontWeights.Bold)
            }
                }
            };
            ReportDataGrid.Columns.Add(ngColumn);

            DataGridTextColumn okColumn = new DataGridTextColumn
            {
                Header = "OK Count",
                Binding = new Binding("OkCount"),
                Width = new DataGridLength(1, DataGridLengthUnitType.Star),
                ElementStyle = new Style(typeof(TextBlock))
                {
                    Setters =
            {
                new Setter(TextBlock.ForegroundProperty, Brushes.Green),
                new Setter(TextBlock.FontWeightProperty, FontWeights.Bold)
            }
                }
            };
            ReportDataGrid.Columns.Add(okColumn);

            DataGridTextColumn totalColumn = new DataGridTextColumn
            {
                Header = "Total Count",
                Binding = new Binding("TotalCount"),
                Width = new DataGridLength(1, DataGridLengthUnitType.Star)
            };
            ReportDataGrid.Columns.Add(totalColumn);
        }




        // For reflection on MeasurementReading
        private string MapParameterToProperty(string parameter) => parameter switch
        {
            "STEP OD1" => nameof(MeasurementReading.StepOd1),
            "STEP RUNOUT-1" => nameof(MeasurementReading.StepRunout1),
            "OD-1" => nameof(MeasurementReading.Od1),
            "RN-1" => nameof(MeasurementReading.Rn1),
            "OD-2" => nameof(MeasurementReading.Od2),
            "RN-2" => nameof(MeasurementReading.Rn2),
            "OD-3" => nameof(MeasurementReading.Od3),
            "RN-3" => nameof(MeasurementReading.Rn3),
            "STEP OD2" => nameof(MeasurementReading.StepOd2),
            "STEP RUNOUT-2" => nameof(MeasurementReading.StepRunout2),
            "ID-1" => nameof(MeasurementReading.Id1),
            "RN-4" => nameof(MeasurementReading.Rn4),
            "ID-2" => nameof(MeasurementReading.Id2),
            "RN-5" => nameof(MeasurementReading.Rn5),
            "OL" => nameof(MeasurementReading.Ol),
            _ => parameter.Replace(" ", "")
        };


        private async void DeleteSelectedRow_Click(object sender, RoutedEventArgs e)
        {
            if (SessionManager.UserType != "Admin")
            {
                MessageBox.Show("Delete access is only available for Admin users.", "Access Denied", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (ReportDataGrid.SelectedItem is ReportTableItem item)
            {
                try
                {
                    // Call your data service method to delete the record
                    await _dataService.DeleteMeasurementReadingAsync(item.Id);

                    // Remove from UI
                    ReportTableItems.Remove(item);
                }
                catch (Exception ex)
                {
                    // Log or show error message if needed
                    MessageBox.Show("Error deleting record: " + ex.Message, "Delete Failed", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }



        private void ReportDataGrid_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Delete)
            {
                // Call the same delete method
                DeleteSelectedRow_Click(sender, e);
                e.Handled = true;
            }
        }


        private void ExportExcel_Click(object sender, RoutedEventArgs e)
        {
            if (!ReportTableItems.Any())
            {
                MessageBox.Show("No data to export.");
                return;
            }

            try
            {
                using (var wb = new XLWorkbook())
                {
                    var ws = wb.Worksheets.Add("Report");

                    // ===============================================================
                    // 🔥 ADD COMPANY LOGO + COMPANY NAME HEADER (Left + Center)
                    // ===============================================================

                    // Logo block (A1:C3)
                    var logoRange = ws.Range("A1:B3");
                    logoRange.Merge();

                    if (!string.IsNullOrEmpty(_companyLogoPath) && File.Exists(_companyLogoPath))
                    {
                        var logoPicture = ws.AddPicture(_companyLogoPath)
                    .MoveTo(ws.Cell("A1"), 4, 4)   // X offset = 5, Y offset = 5
                    .Scale(0.2);                  // Resize logo

                    }

                    // Company Name (Centered between D1:H3)
                    var companyNameRange = ws.Range("C1:G3");
                    companyNameRange.Merge();
                    ws.Cell("C1").Value = _companyName;
                    ws.Cell("C1").Style.Font.Bold = true;
                    ws.Cell("C1").Style.Font.FontSize = 18;
                    ws.Cell("C1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    ws.Cell("C1").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

                    // ===============================================================
                    // 🔥 ORIGINAL LOGIC STARTS — NOTHING CHANGED
                    // ===============================================================

                    var partConfig = _dataService.GetPartConfigByPartNumber(SelectedPartNo ?? ReportTableItems.First().PartNo).ToList();

                    var paramToShort = partConfig
                        .Where(p => !string.IsNullOrWhiteSpace(p.Parameter))
                        .ToDictionary(
                            p => p.Parameter.Trim(),
                            p => string.IsNullOrWhiteSpace(p.ShortName) ? p.Parameter.Trim() : p.ShortName.Trim(),
                            StringComparer.OrdinalIgnoreCase
                        );

                    bool isSingleParameter = SelectedParameter != null && SelectedParameter != "All";
                    List<string> paramList = isSingleParameter
                        ? new List<string> { SelectedParameter }
                        : ReportTableItems.First().Parameters.Keys.ToList();

                    List<string> paramShortList = paramList
                        .Select(p => paramToShort.TryGetValue(p, out var shortName) ? shortName : p)
                        .ToList();

                    var USL = partConfig.Select(p => new ParameterValue { Parameter = p.Parameter, Value = p.Nominal - p.RTolMinus }).ToList();
                    var MEAN = partConfig.Select(p => new ParameterValue { Parameter = p.Parameter, Value = p.Nominal }).ToList();
                    var LSL = partConfig.Select(p => new ParameterValue { Parameter = p.Parameter, Value = p.Nominal + p.RTolPlus }).ToList();

                    // Move your existing headers down because of logo/name
                    int baseHeaderRow = 5;

                    ws.Range("J1:M1").Merge();
                    ws.Cell("J1").Value = "Company Name:";
                    ws.Cell("J1").Style.Font.Bold = true;
                    ws.Cell("J1").Style.Font.FontSize = 14;
                    ws.Cell("J1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                    ws.Range("J2:M2").Merge();
                    ws.Cell("J2").Value = $"Date: {(_selectedDateTimeFrom.HasValue ? _selectedDateTimeTo.Value.ToString("dd-MMM-yyyy") : DateTime.Today.ToString("dd-MMM-yyyy"))}";
                    ws.Cell("J2").Style.Font.Bold = true;
                    ws.Cell("J2").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                    ws.Range("J3:M3").Merge();
                    ws.Cell("J3").Value = $"Part Number: {SelectedPartNo}";
                    ws.Cell("J3").Style.Font.Bold = true;
                    ws.Cell("J3").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                    // Parameter headers
                    for (int i = 0; i < partConfig.Count; i++)
                    {
                        var cell = ws.Cell(baseHeaderRow, 4 + i);
                        cell.Value = paramToShort.TryGetValue(partConfig[i].Parameter, out var shortName)
                            ? shortName
                            : partConfig[i].Parameter;

                        cell.Style.Font.Bold = true;
                        cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#D4E6F1");
                        cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    }

                    var labelFormats = new (string Label, XLColor Color)[]
                    {
                ("USL", XLColor.Red),
                ("MEAN", XLColor.ForestGreen),
                ("LSL", XLColor.Red)
                    };

                    for (int idx = 0; idx < labelFormats.Length; idx++)
                    {
                        var labelCell = ws.Cell(baseHeaderRow + 1 + idx, 3);
                        labelCell.Value = labelFormats[idx].Label;
                        labelCell.Style.Font.Bold = true;
                        labelCell.Style.Font.FontColor = labelFormats[idx].Color;
                        labelCell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                    }

                    ws.Column(3).Width = 15;

                    for (int i = 0; i < partConfig.Count; i++)
                    {
                        ws.Cell(baseHeaderRow + 1, 4 + i).Value = USL[i].Value;
                        ws.Cell(baseHeaderRow + 2, 4 + i).Value = MEAN[i].Value;
                        ws.Cell(baseHeaderRow + 3, 4 + i).Value = LSL[i].Value;

                        ws.Cell(baseHeaderRow + 1, 4 + i).Style.Font.FontColor = XLColor.Red;
                        ws.Cell(baseHeaderRow + 2, 4 + i).Style.Font.FontColor = XLColor.ForestGreen;
                        ws.Cell(baseHeaderRow + 3, 4 + i).Style.Font.FontColor = XLColor.Red;
                    }

                    int lastCol = 4 + partConfig.Count - 1;
                    ws.Range(baseHeaderRow, 3, baseHeaderRow + 3, lastCol)
                          .Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                    int startTableRow = baseHeaderRow + 5;

                    int col = 1;
                    ws.Cell(startTableRow, col++).Value = "S.No";
                    ws.Cell(startTableRow, col++).Value = "Part No";
                    ws.Cell(startTableRow, col++).Value = "Lot No";
                    ws.Cell(startTableRow, col++).Value = "Operator";
                    ws.Cell(startTableRow, col++).Value = "Date";

                    foreach (var p in paramShortList)
                        ws.Cell(startTableRow, col++).Value = p;

                    if (!isSingleParameter)
                        ws.Cell(startTableRow, col++).Value = "Status";

                    int row = startTableRow + 1;

                    foreach (var item in ReportTableItems)
                    {
                        col = 1;
                        ws.Cell(row, col++).Value = item.SerialNo;
                        ws.Cell(row, col++).Value = item.PartNo;
                        ws.Cell(row, col++).Value = item.LotNo;
                        ws.Cell(row, col++).Value = item.Operator;
                        ws.Cell(row, col++).Value = item.Date;

                        foreach (var p in paramList)
                        {
                            var cell = ws.Cell(row, col++);
                            double value = item.Parameters.ContainsKey(p) ? item.Parameters[p] : 0;
                            cell.Value = value;

                            var config = partConfig.FirstOrDefault(pc => pc.Parameter.Equals(p, StringComparison.OrdinalIgnoreCase));
                            if (config != null)
                            {
                                double usl = config.Nominal - config.RTolMinus;
                                double lsl = config.Nominal + config.RTolPlus;

                                if (value < usl || value > lsl)
                                    cell.Style.Font.FontColor = XLColor.Red;
                            }
                        }

                        if (!isSingleParameter)
                            ws.Cell(row, col++).Value = item.Status;

                        row++;
                    }

                    ws.Columns().AdjustToContents();

                    string baseFolder = @"D:\MEPL\Excel Report\Generated Data";
                    if (!Directory.Exists(baseFolder))
                        Directory.CreateDirectory(baseFolder);

                    string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                    string fileName = $"EVMS_Report_{timestamp}.xlsx";
                    string filePath = Path.Combine(baseFolder, fileName);

                    wb.SaveAs(filePath);
                    MessageBox.Show($"Excel exported successfully to: {filePath}");
                    Process.Start(new ProcessStartInfo(filePath) { UseShellExecute = true });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Excel export failed: {ex.Message}");
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

        private void ExportPdf_Click(object sender, RoutedEventArgs e)
        {
            if (!ReportTableItems.Any())
            {
                MessageBox.Show("No data to export.");
                return;
            }

            try
            {
                var document = new PdfDocument();
                var fontOptions = new XPdfFontOptions(PdfFontEncoding.Unicode);

                // Fonts
                var companyFont = new XFont("Arial", 14, XFontStyle.Bold, fontOptions);
                var titleFont = new XFont("Arial", 12, XFontStyle.Bold, fontOptions);
                var headerFont = new XFont("Arial", 7, XFontStyle.Bold, fontOptions);
                var bodyFont = new XFont("Arial", 6, XFontStyle.Regular, fontOptions);
                var infoFont = new XFont("Arial", 7, XFontStyle.Regular, fontOptions);
                var infoFontBold = new XFont("Arial", 7, XFontStyle.Bold, fontOptions);

                double margin = 25;
                double rowHeight = 14;

                // Fetch part config
                var partConfigs = _dataService.GetPartConfigByPartNumber(SelectedPartNo ?? ReportTableItems.First().PartNo).ToList();

                var paramToShort = partConfigs.ToDictionary(
                    p => p.Parameter.Trim(),
                    p => string.IsNullOrWhiteSpace(p.ShortName) ? p.Parameter.Trim() : p.ShortName.Trim(),
                    StringComparer.OrdinalIgnoreCase
                );

                bool isSingleParameter = SelectedParameter != null && SelectedParameter != "All";

                var paramList = isSingleParameter
                    ? new List<string> { SelectedParameter }
                    : ReportTableItems.First().Parameters.Keys.ToList();

                var paramShortList = paramList
                    .Select(p => paramToShort.TryGetValue(p, out var shortName) ? shortName : p)
                    .ToList();

                // FIXED HEADERS
                var fixedHeaders = new List<string> { "S.No", "Part No", "Lot No", "Operator", "Date" };

                // ALL HEADERS
                var allHeaders = new List<string>();
                allHeaders.AddRange(fixedHeaders);
                allHeaders.AddRange(paramShortList);
                if (!isSingleParameter)
                    allHeaders.Add("Status");

                // USL / MEAN / LSL
                var USL = partConfigs.Select(p => p.Nominal + p.RTolPlus).ToList();
                var MEAN = partConfigs.Select(p => p.Nominal).ToList();
                var LSL = partConfigs.Select(p => p.Nominal - p.RTolMinus).ToList();

                PdfPage page = document.AddPage();
                page.Orientation = PdfSharpCore.PageOrientation.Landscape;
                var gfx = XGraphics.FromPdfPage(page);

                double pageWidth = page.Width;
                double pageHeight = page.Height;
                double usableWidth = pageWidth - (2 * margin);

                // ===============================
                //      COLUMN WIDTH SECTION
                // ===============================

                Dictionary<string, double> colWidths = new();

                // FIXED WIDTHS (STABLE)
                colWidths["S.No"] = 40;
                colWidths["Part No"] = 70;
                colWidths["Lot No"] = 55;
                colWidths["Operator"] = 70;
                colWidths["Date"] = 75;

                double fixedWidthSum = colWidths.Sum(c => c.Value);

                // DYNAMIC COLUMNS
                int dynamicCount = paramShortList.Count + (isSingleParameter ? 0 : 1);
                double remainingWidth = usableWidth - fixedWidthSum;
                double dynWidth = remainingWidth / dynamicCount;

                foreach (var p in paramShortList)
                    colWidths[p] = dynWidth;

                if (!isSingleParameter)
                    colWidths["Status"] = dynWidth;

                // Center text safely
                void DrawCentered(XGraphics g, string text, XFont font, XBrush brush, XRect rect)
                {
                    if (string.IsNullOrEmpty(text)) return;
                    var fmt = new XStringFormat
                    {
                        Alignment = XStringAlignment.Center,
                        LineAlignment = XLineAlignment.Center
                    };
                    g.DrawString(text, font, brush, rect, fmt);
                }

                double y = margin;

                // ===============================
                //    DRAW PDF HEADER FUNCTION
                // ===============================

                void DrawHeader(XGraphics g, PdfPage pg, ref double yPos, bool includeMeta)
                {
                    yPos = margin;

                    // 1) LOGO + COMPANY NAME
                    double logoW = 40, logoH = 40;

                    if (!string.IsNullOrEmpty(_companyLogoPath) && File.Exists(_companyLogoPath))
                    {
                        XImage img = XImage.FromFile(_companyLogoPath);
                        g.DrawImage(img, margin, yPos - 15, logoW, logoH);
                        g.DrawString(_companyName, companyFont, XBrushes.Black,
                            new XRect(margin + logoW + 10, yPos - 10, 400, 40), XStringFormats.TopLeft);
                    }
                    else
                    {
                        g.DrawString(_companyName, companyFont, XBrushes.Black,
                            new XRect(margin, yPos - 10, 400, 40), XStringFormats.TopLeft);
                    }

                    yPos += logoH;

                    // 2) TITLE + META DATA
                    if (includeMeta)
                    {
                        var titleRect = new XRect(margin, yPos, usableWidth, 25);
                        g.DrawRectangle(XBrushes.LightGray, titleRect);
                        DrawCentered(g, "MEASUREMENT REPORT", titleFont, XBrushes.Black, titleRect);
                        yPos += 30;

                        double labelW = 70, valW = 120;
                        double x1 = margin, x2 = margin + 250, x3 = margin + 500;

                        DrawCentered(g, "Part No:", infoFontBold, XBrushes.Black, new XRect(x1, yPos, labelW, 10));
                        DrawCentered(g, SelectedPartNo ?? "-", infoFont, XBrushes.Black, new XRect(x1 + labelW, yPos, valW, 10));

                        DrawCentered(g, "Lot No:", infoFontBold, XBrushes.Black, new XRect(x2, yPos, labelW, 10));
                        DrawCentered(g, SelectedLotNo ?? "-", infoFont, XBrushes.Black, new XRect(x2 + labelW, yPos, valW, 10));
                        yPos += 14;

                        DrawCentered(g, "Operator:", infoFontBold, XBrushes.Black, new XRect(x1, yPos, labelW, 10));
                        DrawCentered(g, SelectedOperator ?? "-", infoFont, XBrushes.Black, new XRect(x1 + labelW, yPos, valW, 10));

                        DrawCentered(g, "From:", infoFontBold, XBrushes.Black, new XRect(x2, yPos, labelW, 10));
                        DrawCentered(g, _selectedDateTimeFrom?.ToShortDateString() ?? "-", infoFont, XBrushes.Black, new XRect(x2 + labelW, yPos, valW, 10));

                        DrawCentered(g, "To:", infoFontBold, XBrushes.Black,
                            new XRect(x3, yPos, labelW, 10));
                        DrawCentered(
                                         g,
                                         _selectedDateTimeTo?.ToShortDateString() ?? "-",
                                         infoFont,
                                         XBrushes.Black,
                                         new XRect(x3 + labelW, yPos, valW, 10)
                                     );


                        yPos += 16;
                        g.DrawLine(XPens.Black, margin, yPos, pg.Width - margin, yPos);
                        yPos += 6;
                    }

                    // 3) HEADER ROW
                    double x = margin;
                    foreach (var h in allHeaders)
                    {
                        double w = colWidths[h];
                        var rect = new XRect(x, yPos, w, rowHeight);
                        g.DrawRectangle(XPens.Black, XBrushes.LightGray, rect);
                        DrawCentered(g, h, headerFont, XBrushes.Black, rect);
                        x += w;
                    }
                    yPos += rowHeight;

                    // 4) USL/MEAN/LSL ROWS
                    string[] labels = { "USL", "MEAN", "LSL" };
                    List<List<double>> values = new() { USL, MEAN, LSL };
                    XBrush[] brushes = { XBrushes.Red, XBrushes.Green, XBrushes.Red };

                    foreach (var idx in Enumerable.Range(0, 3))
                    {
                        double xx = margin;

                        // Fixed block label
                        var fixedRect = new XRect(xx, yPos, fixedWidthSum, rowHeight);
                        DrawCentered(g, labels[idx], headerFont, brushes[idx], fixedRect);
                        xx += fixedWidthSum;

                        // Parameter values
                        for (int p = 0; p < paramList.Count; p++)
                        {
                            string fullParam = paramList[p];
                            string shortName = paramShortList[p];

                            double w = colWidths[shortName];

                            double val = 0;
                            var config = partConfigs.FirstOrDefault(c =>
                                c.Parameter.Equals(fullParam, StringComparison.OrdinalIgnoreCase));
                            if (config != null)
                            {
                                int confIndex = partConfigs.IndexOf(config);
                                val = values[idx][confIndex];
                            }

                            var rect = new XRect(xx, yPos, w, rowHeight);
                            DrawCentered(g, val.ToString("0.###"), bodyFont, brushes[idx], rect);
                            xx += w;
                        }

                        // Status column (blank)
                        if (!isSingleParameter)
                        {
                            double w = colWidths["Status"];
                            var rect = new XRect(xx, yPos, w, rowHeight);
                            xx += w;
                        }

                        yPos += rowHeight;
                    }

                    yPos += 5;
                }

                // FIRST PAGE HEADER
                DrawHeader(gfx, page, ref y, true);

                int rowIndex = 0;

                // ===============================
                //           MAIN TABLE
                // ===============================

                foreach (var item in ReportTableItems)
                {
                    double x = margin;
                    XBrush bg = rowIndex % 2 == 0 ? XBrushes.White : XBrushes.Ivory;

                    // fixed values
                    var fixedVals = new List<string>
            {
                item.SerialNo.ToString(),
                item.PartNo,
                item.LotNo,
                item.Operator,
                item.Date
            };

                    for (int i = 0; i < fixedHeaders.Count; i++)
                    {
                        string h = fixedHeaders[i];
                        double w = colWidths[h];

                        var rect = new XRect(x, y, w, rowHeight);
                        gfx.DrawRectangle(XPens.Black, bg, rect);
                        DrawCentered(gfx, fixedVals[i], bodyFont, XBrushes.Black, rect);
                        x += w;
                    }

                    // parameter values
                    for (int i = 0; i < paramList.Count; i++)
                    {
                        string full = paramList[i];
                        string shortName = paramShortList[i];

                        double w = colWidths[shortName];

                        double val = item.Parameters.ContainsKey(full) ? item.Parameters[full] : 0;

                        var config = partConfigs.FirstOrDefault(c =>
                            c.Parameter.Equals(full, StringComparison.OrdinalIgnoreCase));

                        XBrush brush = XBrushes.Black;
                        XBrush back = bg;

                        if (config != null)
                        {
                            double usl = config.Nominal + config.RTolPlus;
                            double lsl = config.Nominal - config.RTolMinus;

                            if (val < lsl || val > usl)
                            {
                                brush = XBrushes.Red;
                                back = XBrushes.MistyRose;
                            }
                        }

                        var rect = new XRect(x, y, w, rowHeight);
                        gfx.DrawRectangle(XPens.Black, back, rect);
                        DrawCentered(gfx, val.ToString("0.###"), bodyFont, brush, rect);

                        x += w;
                    }

                    // Status column
                    if (!isSingleParameter)
                    {
                        double w = colWidths["Status"];
                        var rect = new XRect(x, y, w, rowHeight);
                        gfx.DrawRectangle(XPens.Black, bg, rect);
                        DrawCentered(gfx, item.Status, bodyFont, XBrushes.Black, rect);
                        x += w;
                    }

                    y += rowHeight;
                    rowIndex++;

                    // PAGE BREAK
                    if (y > pageHeight - margin - (rowHeight * 5))
                    {
                        page = document.AddPage();
                        page.Orientation = PdfSharpCore.PageOrientation.Landscape;
                        gfx = XGraphics.FromPdfPage(page);

                        DrawHeader(gfx, page, ref y, false);
                    }
                }

                // SAVE FILE
                string folder = @"D:\MEPL\PDF Report";
                if (!Directory.Exists(folder))
                    Directory.CreateDirectory(folder);

                string fileName = $"EVMS_Report_{DateTime.Now:yyyyMMdd_HHmmss}.pdf";
                string filePath = Path.Combine(folder, fileName);

                document.Save(filePath);
                MessageBox.Show($"PDF exported successfully:\n{filePath}");
                Process.Start(new ProcessStartInfo(filePath) { UseShellExecute = true });
            }
            catch (Exception ex)
            {
                MessageBox.Show("PDF Error: " + ex.Message);
            }
        }



        private void ReportDataGrid_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            if (e.Row.Item is not ReportTableItem item)
                return;

            // Apply cell-level color logic AFTER the row is fully generated
            e.Row.Dispatcher.InvokeAsync(() =>
            {
                foreach (var column in ReportDataGrid.Columns.OfType<DataGridTextColumn>())
                {
                    if (column.Header is string param && item.Parameters.ContainsKey(param))
                    {
                        var cellContent = column.GetCellContent(e.Row) as TextBlock;
                        if (cellContent == null) continue;

                        double value = item.Parameters[param];

                        // Get tolerance from config for this parameter
                        var config = _dataService.GetPartConfig(item.PartNo)
                            .FirstOrDefault(pc => pc.Parameter.Equals(param, StringComparison.OrdinalIgnoreCase));

                        if (config != null)
                        {
                            double usl = config.Nominal + config.RTolPlus;
                            double lsl = config.Nominal - config.RTolMinus;

                            if (value < lsl || value > usl)
                            {
                                // 🔴 Out of range — make red and bold
                                cellContent.Foreground = Brushes.Red;
                                cellContent.FontWeight = FontWeights.Bold;
                            }
                            else
                            {
                                // ✅ Within range — make normal black
                                cellContent.Foreground = Brushes.Black;
                                cellContent.FontWeight = FontWeights.Normal;
                            }
                        }
                        else
                        {
                            // Default color if config missing
                            cellContent.Foreground = Brushes.Black;
                            cellContent.FontWeight = FontWeights.Normal;
                        }
                    }
                }
            },
            System.Windows.Threading.DispatcherPriority.Background);
        }




        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string propertyName) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }

    public class ReportTableItem
    {
        public int Id { get; set; }  // ← Add this line

        public int SerialNo { get; set; }
        public string? PartNo { get; set; }
        public string? LotNo { get; set; }
        public string? Operator { get; set; }
        public string? Date { get; set; }
        public Dictionary<string, double> Parameters { get; set; } = new();
        public string? Status { get; set; }
        public HashSet<string> OutOfRangeParams { get; set; } = new HashSet<string>();

    }

    public class NgOkSummaryItem
    {
        public int Id { get; set; }
        public int SerialNo { get; set; }
        public string? PartNo { get; set; }
        public string? LotNo { get; set; }
        public string? Operator { get; set; }
        public string? Date { get; set; }
        public string? Parameter { get; set; }
        public int NgCount { get; set; }
        public int OkCount { get; set; }
        public int TotalCount => NgCount + OkCount;
        public Dictionary<string, double> Parameters { get; set; } = new Dictionary<string, double>();
    }



}
