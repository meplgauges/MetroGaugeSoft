using ClosedXML.Excel;
using EVMS.Service;
using PdfSharpCore.Drawing;
using PdfSharpCore.Pdf;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using static EVMS.Login_Page;
using static SkiaSharp.HarfBuzz.SKShaper;

namespace EVMS
{
    public partial class Report_View_Page : UserControl, INotifyPropertyChanged
    {
        private readonly DataStorageService _dataService;

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

            LoadDesignOptions();
            LoadActiveParts();
            LoadParameters();

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
                string partFilter = SelectedPartNo;
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

            string partFilter = SelectedPartNo == "All" ? null : SelectedPartNo;
            string lotFilter = SelectedLotNo == "All" ? null : SelectedLotNo;
            string operatorFilter = SelectedOperator == "All" ? null : SelectedOperator;

            var results = await _dataService.GetMeasurementReadingsAsync(
                partFilter, lotFilter, operatorFilter, _selectedDateTimeFrom, _selectedDateTimeTo);

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
                        var propName = MapParameterToColumn(param);
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

            string partFilter = SelectedPartNo == "All" ? null : SelectedPartNo;
            string lotFilter = SelectedLotNo == "All" ? null : SelectedLotNo;
            string operatorFilter = SelectedOperator == "All" ? null : SelectedOperator;

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

                string propName = MapParameterToColumn(param);

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

            // Keep fixed columns and remove existing parameter columns
            while (ReportDataGrid.Columns.Count > 5)
                ReportDataGrid.Columns.RemoveAt(5);

            foreach (var param in parameters)
            {
                DataGridTextColumn column = new()
                {
                    Header = param,
                    Binding = new System.Windows.Data.Binding($"Parameters[{param}]"),
                    Width = 81
                };
                ReportDataGrid.Columns.Add(column);
            }

            // Add Status column always at the end
            ReportDataGrid.Columns.Add(new DataGridTextColumn
            {
                Header = "Status",
                Binding = new System.Windows.Data.Binding("Status"),
                Width = 65
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


        private string MapParameterToColumn(string parameter)
        {
            return parameter switch
            {
                "Overall Length" => "OL",
                "Datum to End" => "DE",
                "Head Diameter" => "HD",
                "Groove Position" => "GP",
                "Stem Dia Near Groove" => "STDG",
                "Stem Dia Near Undercut" => "STDU",
                "Groove Diameter" => "GIR_DIA",
                "Straightness" => "STN",
                "Ovality SDG" => "Ovality_SDG",
                "Ovality SDU" => "Ovality_SDU",
                "Ovality Head" => "Ovality_Head",
                "Stem Taper" => "Stem_Taper",
                "End Face Runout" => "EFRO",
                "Face Runout" => "Face_Runout",
                "Seat Height" => "SH",
                "Datum to Groove" => "DG",
                "Seat Runout" => "S_RO",
                _ => parameter?.Replace(" ", "") ?? string.Empty
            };
        }

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

                    // 🔹 Retrieve part config matching selected part or first item’s part
                    var partConfig = _dataService.GetPartConfigByPartNumber(SelectedPartNo ?? ReportTableItems.First().PartNo).ToList();

                    // 🔹 Create ShortName mapping
                    var paramToShort = partConfig
                        .Where(p => !string.IsNullOrWhiteSpace(p.Parameter))
                        .ToDictionary(
                            p => p.Parameter.Trim(),
                            p => string.IsNullOrWhiteSpace(p.ShortName) ? p.Parameter.Trim() : p.ShortName.Trim(),
                            StringComparer.OrdinalIgnoreCase
                        );

                    // 🔹 Determine if single parameter
                    bool isSingleParameter = SelectedParameter != null && SelectedParameter != "All";
                    List<string> paramList = isSingleParameter
                        ? new List<string> { SelectedParameter }
                        : ReportTableItems.First().Parameters.Keys.ToList();

                    // 🔹 Convert full parameter names to short names
                    List<string> paramShortList = paramList
                        .Select(p => paramToShort.TryGetValue(p, out var shortName) ? shortName : p)
                        .ToList();

                    // 🔹 Calculate USL, MEAN, LSL
                    var USL = partConfig.Select(p => new ParameterValue { Parameter = p.Parameter, Value = p.Nominal - p.RTolMinus }).ToList();
                    var MEAN = partConfig.Select(p => new ParameterValue { Parameter = p.Parameter, Value = p.Nominal }).ToList();
                    var LSL = partConfig.Select(p => new ParameterValue { Parameter = p.Parameter, Value = p.Nominal + p.RTolPlus }).ToList();

                    // 🔹 Company/Part info header
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

                    // 🔹 Parameter header row (Row 5, start at column D)
                    for (int i = 0; i < partConfig.Count; i++)
                    {
                        var cell = ws.Cell(5, 4 + i);
                        cell.Value = paramToShort.TryGetValue(partConfig[i].Parameter, out var shortName)
                            ? shortName
                            : partConfig[i].Parameter;

                        cell.Style.Font.Bold = true;
                        cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#D4E6F1");
                        cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    }

                    // 🔹 Label setup for USL, MEAN, LSL (Rows 6-8)
                    var labelFormats = new (string Label, XLColor Color)[]
                    {
                ("USL", XLColor.Red),
                ("MEAN", XLColor.ForestGreen),
                ("LSL", XLColor.Red)
                    };

                    for (int idx = 0; idx < labelFormats.Length; idx++)
                    {
                        var labelCell = ws.Cell(6 + idx, 3);
                        labelCell.Value = labelFormats[idx].Label;
                        labelCell.Style.Font.Bold = true;
                        labelCell.Style.Font.FontColor = labelFormats[idx].Color;
                        labelCell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                    }

                    ws.Column(3).Width = 15;

                    // 🔹 Fill USL/MEAN/LSL values
                    for (int i = 0; i < partConfig.Count; i++)
                    {
                        var uslCell = ws.Cell(6, 4 + i);
                        uslCell.Value = USL[i].Value;
                        uslCell.Style.Font.Bold = true;
                        uslCell.Style.Font.FontColor = XLColor.Red;
                        uslCell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        uslCell.Style.NumberFormat.Format = "0.000";

                        var meanCell = ws.Cell(7, 4 + i);
                        meanCell.Value = MEAN[i].Value;
                        meanCell.Style.Font.Bold = true;
                        meanCell.Style.Font.FontColor = XLColor.ForestGreen;
                        meanCell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        meanCell.Style.NumberFormat.Format = "0.000";

                        var lslCell = ws.Cell(8, 4 + i);
                        lslCell.Value = LSL[i].Value;
                        lslCell.Style.Font.Bold = true;
                        lslCell.Style.Font.FontColor = XLColor.Red;
                        lslCell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        lslCell.Style.NumberFormat.Format = "0.000";
                    }

                    int lastCol = 4 + partConfig.Count - 1;
                    var range = ws.Range(5, 3, 8, lastCol);
                    range.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    range.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

                    // 🔹 Table headers start at row 10
                    int startTableRow = 10;
                    int col = 1;
                    ws.Cell(startTableRow, col++).Value = "S.No";
                    ws.Cell(startTableRow, col++).Value = "Part No";
                    ws.Cell(startTableRow, col++).Value = "Lot No";
                    ws.Cell(startTableRow, col++).Value = "Operator";
                    ws.Cell(startTableRow, col++).Value = "Date";

                    foreach (var p in paramShortList)
                        ws.Cell(startTableRow, col++).Value = p;

                    // 🔸 Add Status only if NOT single parameter
                    if (!isSingleParameter)
                        ws.Cell(startTableRow, col++).Value = "Status";

                    // 🔹 Fill table data
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

                            // Find matching part config
                            var config = partConfig.FirstOrDefault(pc =>
                                pc.Parameter.Equals(p, StringComparison.OrdinalIgnoreCase));

                            if (config != null)
                            {
                                double usl = config.Nominal - config.RTolMinus;
                                double lsl = config.Nominal + config.RTolPlus;

                                // Check if out of range
                                if (value < usl || value > lsl)
                                {
                                    cell.Style.Font.FontColor = XLColor.Red;
                                }
                                else
                                {
                                    cell.Style.Font.FontColor = XLColor.Black;
                                }
                            }

                            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            cell.Style.NumberFormat.Format = "0.000";
                        }


                        // 🔸 Skip Status if single parameter
                        if (!isSingleParameter)
                            ws.Cell(row, col++).Value = item.Status;

                        row++;
                    }

                    // 🔹 Style header row
                    var headerRange = ws.Range(startTableRow, 1, startTableRow, col - 1);
                    headerRange.Style.Font.Bold = true;
                    headerRange.Style.Fill.BackgroundColor = XLColor.LightGray;
                    headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    headerRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    headerRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

                    ws.Columns().AdjustToContents();

                    // 🔹 Save file
                    string baseFolder = @"E:\MEPL\Excel Report\Generated Data";
                    if (!Directory.Exists(baseFolder))
                        Directory.CreateDirectory(baseFolder);

                    string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                    string fileName = $"EVMS_Report_{timestamp}.xlsx";
                    string filePath = Path.Combine(baseFolder, fileName);

                    wb.SaveAs(filePath);
                    MessageBox.Show($"✅ Excel exported successfully to: {filePath}");
                    Process.Start(new ProcessStartInfo(filePath) { UseShellExecute = true });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"❌ Excel export failed: {ex.Message}");
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

                var fixedHeaders = new List<string> { "S.No", "Part No", "Lot No", "Operator", "Date" };
                var allHeaders = new List<string>(fixedHeaders);
                allHeaders.AddRange(paramShortList);

                if (!isSingleParameter)
                    allHeaders.Add("Status");

                var USL = partConfigs.Select(p => p.Nominal - p.RTolMinus).ToList();
                var MEAN = partConfigs.Select(p => p.Nominal).ToList();
                var LSL = partConfigs.Select(p => p.Nominal + p.RTolPlus).ToList();

                PdfPage page = document.AddPage();
                page.Orientation = PdfSharpCore.PageOrientation.Landscape;
                var gfx = XGraphics.FromPdfPage(page);

                // 🔹 Load Logo Once
                string logoPath = @"D:\Project\MetroGaugeSoft\NITTAN 0.1\EVMS\Resource\0987.png";
                XImage logo = XImage.FromFile(logoPath);

                double pageWidth = page.Width;
                double pageHeight = page.Height;
                double usableWidth = pageWidth - (2 * margin);

                var dynamicWidths = new Dictionary<string, double>();
                int totalCols = allHeaders.Count;

                foreach (var h in allHeaders)
                {
                    if (h == "Date") dynamicWidths[h] = usableWidth * 0.10;
                    else if (h == "Lot No") dynamicWidths[h] = usableWidth * 0.06;
                    else dynamicWidths[h] = usableWidth * (0.84 / (totalCols - 2));
                }

                void DrawStringCentered(XGraphics g, string text, XFont font, XBrush brush, XRect rect)
                {
                    if (string.IsNullOrEmpty(text)) return;
                    var measured = g.MeasureString(text, font);
                    double dx = rect.X + Math.Max(0, (rect.Width - measured.Width) / 2.0);
                    double dy = rect.Y + Math.Max(0, (rect.Height - measured.Height) / 2.0);
                    g.DrawString(text, font, brush, new XRect(dx, dy, measured.Width, measured.Height), XStringFormats.TopLeft);
                }

                double y = margin;

                // DRAW HEADER FUNCTION
                void DrawHeader(XGraphics g, PdfPage pg, ref double yPos, bool includeMeta)
                {
                    yPos = margin;

                    if (includeMeta)
                    {
                        var titleRect = new XRect(margin, yPos, pg.Width - 2 * margin, 25);
                        g.DrawRectangle(XBrushes.LightGray, titleRect);
                        DrawStringCentered(g, "MEASUREMENT REPORT", titleFont, XBrushes.Black, titleRect);
                        yPos += 30;

                        double labelWidth = 70, valueWidth = 150, x1 = margin, x2 = margin + 200, x3 = margin + 400;

                        DrawStringCentered(g, "Part No:", infoFontBold, XBrushes.Black, new XRect(x1, yPos, labelWidth, 10));
                        DrawStringCentered(g, SelectedPartNo ?? "-", infoFont, XBrushes.Black, new XRect(x1 + labelWidth, yPos, valueWidth, 10));

                        DrawStringCentered(g, "Lot No:", infoFontBold, XBrushes.Black, new XRect(x2, yPos, labelWidth, 10));
                        DrawStringCentered(g, SelectedLotNo ?? "-", infoFont, XBrushes.Black, new XRect(x2 + labelWidth, yPos, valueWidth, 10));
                        yPos += 14;

                        DrawStringCentered(g, "Operator:", infoFontBold, XBrushes.Black, new XRect(x1, yPos, labelWidth, 10));
                        DrawStringCentered(g, SelectedOperator ?? "-", infoFont, XBrushes.Black, new XRect(x1 + labelWidth, yPos, valueWidth, 10));

                        DrawStringCentered(g, "From:", infoFontBold, XBrushes.Black, new XRect(x2, yPos, labelWidth, 10));
                        DrawStringCentered(g, _selectedDateTimeFrom?.ToShortDateString() ?? "-", infoFont, XBrushes.Black, new XRect(x2 + labelWidth, yPos, valueWidth, 10));

                        DrawStringCentered(g, "To:", infoFontBold, XBrushes.Black, new XRect(x3, yPos, labelWidth, 10));
                        DrawStringCentered(g, _selectedDateTimeTo?.ToShortDateString() ?? "-", infoFont, XBrushes.Black, new XRect(x3 + labelWidth, yPos, valueWidth, 10));
                        yPos += 16;

                        g.DrawLine(XPens.Black, margin, yPos, pg.Width - margin, yPos);
                        yPos += 6;
                    }

                    double x = margin;
                    foreach (var header in allHeaders)
                    {
                        var rect = new XRect(x, yPos, dynamicWidths[header], rowHeight);
                        g.DrawRectangle(XPens.Black, XBrushes.LightGray, rect);
                        DrawStringCentered(g, header, headerFont, XBrushes.Black, rect);
                        x += dynamicWidths[header];
                    }

                    yPos += rowHeight;

                    string[] summaryLabels = { "USL", "MEAN", "LSL" };
                    List<List<double>> summaryValues = new() { USL, MEAN, LSL };
                    XBrush[] brushes = { XBrushes.Red, XBrushes.ForestGreen, XBrushes.Red };

                    for (int s = 0; s < 3; s++)
                    {
                        x = margin;
                        DrawStringCentered(g, summaryLabels[s], headerFont, brushes[s],
                            new XRect(x, yPos, dynamicWidths["S.No"] * fixedHeaders.Count, rowHeight));

                        foreach (var fh in fixedHeaders)
                            x += dynamicWidths[fh];

                        for (int i = 0; i < paramList.Count; i++)
                        {
                            string param = paramList[i];
                            var config = partConfigs.FirstOrDefault(pc => pc.Parameter.Equals(param, StringComparison.OrdinalIgnoreCase));

                            double val = 0;
                            if (config != null)
                            {
                                int index = partConfigs.IndexOf(config);
                                if (index >= 0 && index < summaryValues[s].Count)
                                    val = summaryValues[s][index];
                            }

                            string shortHeader = paramToShort.TryGetValue(param, out var shortName) ? shortName : param;
                            double colWidth = dynamicWidths.ContainsKey(shortHeader) ? dynamicWidths[shortHeader] : (usableWidth / Math.Max(1, paramList.Count));

                            var rect = new XRect(x, yPos, colWidth, rowHeight);
                            DrawStringCentered(g, val.ToString("0.###"), bodyFont, brushes[s], rect);
                            x += colWidth;
                        }

                        yPos += rowHeight;
                    }

                    yPos += 5;

                    // 🔹 Draw small logo bottom-right
                    double logoWidth = 40;       // small size
                    double logoHeight = 40;
                    double logoX = pg.Width - margin - logoWidth;
                    double logoY = pg.Height - margin - logoHeight;

                    g.DrawImage(logo, logoX, logoY, logoWidth, logoHeight);
                }

                // First page header + logo
                DrawHeader(gfx, page, ref y, true);

                int rowCounter = 0;

                foreach (var item in ReportTableItems)
                {
                    double x = margin;
                    XBrush rowBrush = rowCounter % 2 == 0 ? XBrushes.White : XBrushes.Ivory;

                    var fixedValues = new List<string>
            {
                item.SerialNo.ToString(),
                item.PartNo,
                item.LotNo,
                item.Operator,
                item.Date
            };

                    for (int i = 0; i < fixedHeaders.Count; i++)
                    {
                        var rect = new XRect(x, y, dynamicWidths[fixedHeaders[i]], rowHeight);
                        gfx.DrawRectangle(XPens.Black, rowBrush, rect);
                        DrawStringCentered(gfx, fixedValues[i], bodyFont, XBrushes.Black, rect);
                        x += dynamicWidths[fixedHeaders[i]];
                    }

                    for (int i = 0; i < paramList.Count; i++)
                    {
                        string fullParam = paramList[i];
                        string shortName = paramShortList[i];

                        double colWidth = dynamicWidths.ContainsKey(shortName)
      ? dynamicWidths[shortName]
      : (usableWidth / Math.Max(1, paramList.Count));


                        var rect = new XRect(x, y, colWidth, rowHeight);

                        XBrush backgroundBrush = rowBrush;
                        XBrush textBrush = XBrushes.Black;

                        double val = item.Parameters.ContainsKey(fullParam) ? item.Parameters[fullParam] : 0;

                        var config = partConfigs.FirstOrDefault(pc =>
                            pc.Parameter.Equals(fullParam, StringComparison.OrdinalIgnoreCase));

                        if (config != null)
                        {
                            double usl = config.Nominal + config.RTolPlus;
                            double lsl = config.Nominal - config.RTolMinus;

                            if (val < lsl || val > usl)
                            {
                                textBrush = XBrushes.Red;
                                backgroundBrush = XBrushes.MistyRose;
                            }
                        }

                        gfx.DrawRectangle(XPens.Black, backgroundBrush, rect);
                        DrawStringCentered(gfx, val.ToString("0.###"), bodyFont, textBrush, rect);

                        x += colWidth;
                    }

                    if (!isSingleParameter)
                    {
                        var statusRect = new XRect(x, y, dynamicWidths["Status"], rowHeight);
                        gfx.DrawRectangle(XPens.Black, rowBrush, statusRect);
                        DrawStringCentered(gfx, item.Status, bodyFont, XBrushes.Black, statusRect);
                        x += dynamicWidths["Status"];
                    }

                    y += rowHeight;
                    rowCounter++;

                    if (y > pageHeight - margin - (rowHeight * 4))
                    {
                        page = document.AddPage();
                        page.Orientation = PdfSharpCore.PageOrientation.Landscape;
                        gfx = XGraphics.FromPdfPage(page);

                        // NEW PAGE — header + logo
                        DrawHeader(gfx, page, ref y, false);
                    }
                }

                string baseFolder = @"D:\MEPL\PDF Report";
                if (!Directory.Exists(baseFolder))
                    Directory.CreateDirectory(baseFolder);

                string fileName = $"EVMS_Report_{DateTime.Now:yyyyMMdd_HHmmss}.pdf";
                string filePath = Path.Combine(baseFolder, fileName);

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
        public string PartNo { get; set; }
        public string LotNo { get; set; }
        public string Operator { get; set; }
        public string Date { get; set; }
        public Dictionary<string, double> Parameters { get; set; } = new();
        public string Status { get; set; }
        public HashSet<string> OutOfRangeParams { get; set; } = new HashSet<string>();

    }

    public class NgOkSummaryItem
    {
        public int Id { get; set; }
        public int SerialNo { get; set; }
        public string PartNo { get; set; }
        public string LotNo { get; set; }
        public string Operator { get; set; }
        public string Date { get; set; }
        public string Parameter { get; set; }
        public int NgCount { get; set; }
        public int OkCount { get; set; }
        public int TotalCount => NgCount + OkCount;
        public Dictionary<string, double> Parameters { get; set; } = new Dictionary<string, double>();
    }



}
