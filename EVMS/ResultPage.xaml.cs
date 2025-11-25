
using ClosedXML.Excel;
using EVMS.Service;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using static EVMS.Service.MasterService;

namespace EVMS
{
    public partial class ResultPage : UserControl
    {
        public event Action<string>? StatusMessageChanged;

        private Dictionary<string, UserControl> _progressBarControls = new Dictionary<string, UserControl>();
        public event PropertyChangedEventHandler? PropertyChanged;
        private Dictionary<string, string> fullToShortMap;
        private ProcedureMode _currentMode;
        public event EventHandler Closed;
        private DispatcherTimer shiftTimer;





        private bool useFirstDesign = true;
        private bool _showLeft = false;
        private List<PartReadingDataModel> parameterData;
        private string activePartNumber = string.Empty;
        private DataTable _measurementDataTable;
        private const int MaxRows = 10; // Show only last 10 cycles


        private PlcProbeService plcProbeService;
        private DataStorageService dataStorageService;
        private MasterService _masterService;


        public string PartNo { get; set; }
        public string LotNo { get; set; }
        public string OperatorID { get; set; }

        private string _model;
        private string _lotNo;
        private string _userId;


        public ResultPage(string model, string lotNo, string userId)
        {
            InitializeComponent();


            _model = model;
            _lotNo = lotNo;
            _userId = userId;

            SetData(_model, _lotNo, _userId);
            this.Loaded += AutoManual_Loaded;

            this.Loaded += ResultPage_Loaded;
            this.Unloaded += ResultPage_Unloaded;
            this.Loaded += UserControl_Loaded;


            dataStorageService = new DataStorageService();
            _masterService = new MasterService();
            plcProbeService = new PlcProbeService();
            StartShiftTimer();

            _masterService.CalculatedValuesWithStatusReady += MasterService_CalculatedValuesWithStatusReady;
            //_masterService.MeasurementCycleReset += OnCycleReset;

            ValveReadingsGrid.PreviewKeyDown += ValveReadingsGrid_PreviewKeyDown;


            _masterService.StatusMessageUpdated += (message) =>
            {
                StatusMessageChanged?.Invoke(message);
            };

            // Set DataContext for data binding
            this.DataContext = this;


        }
        //public class ParameterResult
        //{
        //    public double Value { get; set; }
        //    public bool IsOk { get; set; }
        //}

        public class MeasurementDataModel
        {
            public string PartNo { get; set; }
            public string LotNo { get; set; }
            public string Operator { get; set; }
            public string Date { get; set; }
            public Dictionary<string, double> Parameters { get; set; } = new Dictionary<string, double>();
            public string Status { get; set; } = string.Empty;
        }


        private void ValveReadingsGrid_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Delete)
            {
                HandleDeleteLatestMeasurementRowAndCounts();
                e.Handled = true;
            }
        }

        private void OnCycleReset()
        {
            // Ensure this runs on the UI thread
            Dispatcher.Invoke(() =>
            {
                ResetMeasurementFieldsAndProgressBars();
            });
        }

        private async Task LoadAndDisplayInspectionDataAsync()
        {
            var existingRecord = await dataStorageService.SelectInspectionDataAsync(_model, _lotNo, _userId);
            // dataStorageService.GetMasterReadingByPart(_model);

            if (existingRecord != null)
            {
                txtModel.Text = existingRecord.PartNo;
                txtLotNo.Text = existingRecord.LotNo;
                txtUserId.Text = existingRecord.OperatorID;
                txtInspectionQty.Text = existingRecord.InspectionQty.ToString();
                txtOkCount.Text = existingRecord.OkCount.ToString();

                int ngCount = existingRecord.InspectionQty - existingRecord.OkCount;
                txtNgCount.Text = ngCount.ToString();
            }
            else
            {
                // Insert a new record with zero counts if not exists
                await dataStorageService.InsertInspectionDataAsync(_model, _lotNo, _userId);

                // Display initial zero data
                txtModel.Text = _model;
                txtLotNo.Text = _lotNo;
                txtUserId.Text = _userId;
                txtInspectionQty.Text = "0";
                txtOkCount.Text = "0";
                txtNgCount.Text = "0";
            }
        }

        private async Task UpdateInspectionDataAsync()
        {
            if (int.TryParse(txtInspectionQty.Text, out int inspectionQty)
                && int.TryParse(txtOkCount.Text, out int okCount))
            {
                await dataStorageService.UpdateInspectionCountsAsync(_model, _lotNo, _userId, inspectionQty, okCount);
            }
            else
            {
                // Handle parse error if necessary
            }
        }


        private async void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            await LoadAndDisplayInspectionDataAsync();
        }
        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public void SetData(string model, string lotNo, string userId)
        {
            txtModel.Text = model;
            txtLotNo.Text = lotNo;
            txtUserId.Text = userId;
        }

        private void NotifyStatus(string message)
        {
            MainWindow.ShowStatusMessage(message);
        }



        private void MasterService_CalculatedValuesWithStatusReady(object? sender, Dictionary<string, ParameterResult> resultsWithStatus)
        {
            Dispatcher.Invoke(() =>
            {
                // --- Update UI ---
                UpdateProgressBarsWithStatus(resultsWithStatus);
                UpdateMeasurementFields(resultsWithStatus);

                var latestValues = resultsWithStatus.ToDictionary(kvp => kvp.Key, kvp => kvp.Value.Value);
                LoadDataGrid(latestValues, resultsWithStatus);
                UpdateInspectionCounts(resultsWithStatus);

                string status = resultsWithStatus.All(r => r.Value.IsOk) ? "OK" : "NG";
                // Build MeasurementDataModel from current data
                var measurement = new MeasurementDataModel
                {
                    PartNo = _model,
                    LotNo = _lotNo,
                    Operator = _userId,
                    Date = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                    Status = status,
                    Parameters = resultsWithStatus.ToDictionary(kvp => kvp.Key, kvp => (double)kvp.Value.Value)
                };

                // Export to Excel asynchronously without blocking UI
                if (_currentMode == ProcedureMode.Measurement && status == "OK")
                {
                    _ = Task.Run(() =>
                    {
                        ExportOrAppendMeasurementToExcel(measurement);
                    });
                }


                // --- Save data asynchronously ---
                _ = Task.Run(async () =>
                {
                    try
                    {
                        // ✅ Get only enabled parameters directly from DB (already filtered)
                        var enabledParams = dataStorageService
                            .GetPartConfig(_model) // query already has IsEnabled = 1
                            .Select(p => p.Parameter.ToLower())
                            .ToHashSet();

                        // ✅ Safe getter: returns 0 for missing or disabled parameters
                        float GetValue(string key)
                        {
                            string lowerKey = key.ToLower();
                            return enabledParams.Contains(lowerKey)
                                ? resultsWithStatus.TryGetValue(key, out var param) ? (float)param.Value : 0f
                                : 0f;
                        }

                        if (_currentMode == ProcedureMode.MasterInspection)
                        {
                            await dataStorageService.InsertMasterInspectionAsync(
                                _model, _userId, _lotNo,
                                GetValue("Overall Length"), GetValue("Datum to End"), GetValue("Head Diameter"),
                                GetValue("Groove Position"), GetValue("Stem Dia Near Groove"), GetValue("Stem Dia Near Undercut"),
                                GetValue("Groove Diameter"), GetValue("Straightness"), GetValue("Ovality SDG"),
                                GetValue("Ovality SDU"), GetValue("Ovality Head"), GetValue("Stem Taper"),
                                GetValue("End Face Runout"), GetValue("Face Runout"), GetValue("Seat Height"),
                                GetValue("Seat Runout"), GetValue("Datum to Groove"), status);
                        }
                        else if (_currentMode == ProcedureMode.Measurement)
                        {
                            await dataStorageService.InsertMeasurementReadingAsync(
                                _model, _userId, _lotNo,
                                GetValue("Overall Length"), GetValue("Datum to End"), GetValue("Head Diameter"),
                                GetValue("Groove Position"), GetValue("Stem Dia Near Groove"), GetValue("Stem Dia Near Undercut"),
                                GetValue("Groove Diameter"), GetValue("Straightness"), GetValue("Ovality SDG"),
                                GetValue("Ovality SDU"), GetValue("Ovality Head"), GetValue("Stem Taper"),
                                GetValue("End Face Runout"), GetValue("Face Runout"), GetValue("Seat Height"),
                                GetValue("Seat Runout"), GetValue("Datum to Groove"), status);
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"❌ Error saving results: {ex.Message}");
                    }
                });

                // --- Reset UI after delay ---
                _ = Task.Run(async () =>
                {
                    await Task.Delay(1500);
                    Dispatcher.Invoke(ResetMeasurementFieldsAndProgressBars);
                });
            });
        }



        private void ExportOrAppendMeasurementToExcel(MeasurementDataModel measurement)
        {

            if (_currentMode != ProcedureMode.Measurement)
                return;

                string baseFolder = @"E:\MEPL\Excel Report\Ok Parts";

            //string baseFolder = Path.Combine(
            //        Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            //        "MEPL", "Excel Report");
            Directory.CreateDirectory(baseFolder);
            if (!Directory.Exists(baseFolder))
                Directory.CreateDirectory(baseFolder);

            // ✅ Sanitize file name parts (remove invalid characters)
            string safePart = string.Concat((measurement.PartNo ?? "UnknownPart").Split(Path.GetInvalidFileNameChars())).Trim();
            string safeLot = string.Concat((measurement.LotNo ?? "UnknownLot").Split(Path.GetInvalidFileNameChars())).Trim();
            string safeOperator = string.Concat((measurement.Operator ?? "UnknownOperator").Split(Path.GetInvalidFileNameChars())).Trim();

            // ✅ Include PartNo, LotNo, Operator in file name
            string fileName = $"EVMS_Report_{safePart}_{safeLot}_{safeOperator}.xlsx";
            string filePath = Path.Combine(baseFolder, fileName);

            bool fileExists = File.Exists(filePath);

            using (var wb = fileExists ? new XLWorkbook(filePath) : new XLWorkbook())
            {
                var ws = wb.Worksheets.Contains("Measurement")
                    ? wb.Worksheet("Measurement")
                    : wb.AddWorksheet("Measurement");

                var partConfig = dataStorageService.GetPartConfigByPartNumber(measurement.PartNo).ToList();

                // 🔹 Company / Part info header (top right)
                ws.Range("J1:M1").Merge();
                ws.Cell("J1").Value = "Company Name:";
                ws.Cell("J1").Style.Font.Bold = true;
                ws.Cell("J1").Style.Font.FontSize = 14;
                ws.Cell("J1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                ws.Range("J2:M2").Merge();
                ws.Cell("J2").Value = $"Date: {DateTime.Today:dd-MMM-yyyy}";
                ws.Cell("J2").Style.Font.Bold = true;
                ws.Cell("J2").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                ws.Range("J3:M3").Merge();
                ws.Cell("J3").Value = $"Part Number: {measurement.PartNo}";
                ws.Cell("J3").Style.Font.Bold = true;
                ws.Cell("J3").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                // 🔹 Calculate USL / MEAN / LSL
                var USL = partConfig.Select(p => new ParameterValue { Parameter = p.Parameter, Value = p.Nominal - p.RTolMinus }).ToList();
                var MEAN = partConfig.Select(p => new ParameterValue { Parameter = p.Parameter, Value = p.Nominal }).ToList();
                var LSL = partConfig.Select(p => new ParameterValue { Parameter = p.Parameter, Value = p.Nominal + p.RTolPlus }).ToList();

                // 🔹 If file new → add layout and headers
                if (!fileExists)
                {
                    int headerStartCol = 4; // Column D
                    ws.Cell(5, 3).Value = "Parameter";
                    ws.Cell(5, 3).Style.Font.Bold = true;
                    ws.Cell(5, 3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                    // Parameter headers
                    for (int i = 0; i < partConfig.Count; i++)
                    {
                        var cell = ws.Cell(5, headerStartCol + i);
                        string shortName = string.IsNullOrWhiteSpace(partConfig[i].ShortName)
                            ? partConfig[i].Parameter
                            : partConfig[i].ShortName;
                        cell.Value = shortName;
                        cell.Style.Font.Bold = true;
                        cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#D4E6F1");
                        cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    }

                    // USL, MEAN, LSL rows
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
                    }

                    // Fill USL/MEAN/LSL values
                    for (int i = 0; i < partConfig.Count; i++)
                    {
                        ws.Cell(6, headerStartCol + i).Value = USL[i].Value;
                        ws.Cell(6, headerStartCol + i).Style.Font.FontColor = XLColor.Red;

                        ws.Cell(7, headerStartCol + i).Value = MEAN[i].Value;
                        ws.Cell(7, headerStartCol + i).Style.Font.FontColor = XLColor.ForestGreen;

                        ws.Cell(8, headerStartCol + i).Value = LSL[i].Value;
                        ws.Cell(8, headerStartCol + i).Style.Font.FontColor = XLColor.Red;

                        for (int r = 6; r <= 8; r++)
                        {
                            ws.Cell(r, headerStartCol + i).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            ws.Cell(r, headerStartCol + i).Style.NumberFormat.Format = "0.000";
                        }
                    }

                    int lastCol = headerStartCol + partConfig.Count - 1;
                    var borderRange = ws.Range(5, 3, 8, lastCol);
                    borderRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    borderRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                     
                    // Measurement table headers
                    int startTableRow = 10;
                    int col = 1;
                    ws.Cell(startTableRow, col++).Value = "S.No";
                    ws.Cell(startTableRow, col++).Value = "Part No";
                    ws.Cell(startTableRow, col++).Value = "Lot No";
                    ws.Cell(startTableRow, col++).Value = "Operator";
                    ws.Cell(startTableRow, col++).Value = "Date";

                    foreach (var p in partConfig)
                        ws.Cell(startTableRow, col++).Value = string.IsNullOrWhiteSpace(p.ShortName)
                            ? p.Parameter
                            : p.ShortName;

                    var headerRange = ws.Range(startTableRow, 1, startTableRow, col - 1);
                    headerRange.Style.Font.Bold = true;
                    headerRange.Style.Fill.BackgroundColor = XLColor.LightGray;
                    headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    headerRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    headerRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                }

                // 🔹 Find next row for new measurement
                int lastRow = ws.LastRowUsed()?.RowNumber() ?? 10;
                int newRow = lastRow + 1;

                // Continue serial numbering
                int nextSerial = 1;
                if (lastRow > 10)
                {
                    var lastSerialCell = ws.Cell(lastRow, 1);
                    if (lastSerialCell.TryGetValue<int>(out int val))
                        nextSerial = val + 1;
                }

                // Write new measurement
                int cIndex = 1;
                ws.Cell(newRow, cIndex++).Value = nextSerial;
                ws.Cell(newRow, cIndex++).Value = measurement.PartNo;
                ws.Cell(newRow, cIndex++).Value = measurement.LotNo;
                ws.Cell(newRow, cIndex++).Value = measurement.Operator;
                ws.Cell(newRow, cIndex++).Value = measurement.Date;

                foreach (var param in partConfig)
                {
                    var cell = ws.Cell(newRow, cIndex++);
                    if (measurement.Parameters.TryGetValue(param.Parameter, out double val))
                    {
                        cell.Value = val;
                        double usl = param.Nominal - param.RTolMinus;
                        double lsl = param.Nominal + param.RTolPlus;
                        cell.Style.Font.FontColor = (val < usl || val > lsl)
                            ? XLColor.Red
                            : XLColor.Black;
                    }
                    else
                    {
                        cell.Value = "";
                    }

                    cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    cell.Style.NumberFormat.Format = "0.000";
                }

                ws.Columns().AdjustToContents();
                wb.SaveAs(filePath);
            }

            // 🟢 No UI message here — runs silently in background
        }




        private void UpdateInspectionCounts(Dictionary<string, ParameterResult> resultsWithStatus)
        {
            if (_currentMode != ProcedureMode.Measurement)
                return;
            
               // Skip incrementing counts during Master Inspection
            Dispatcher.Invoke(() =>
            {
                // Increment InspectionQty correctly
                int inspected;
                if (!int.TryParse(txtInspectionQty.Text, out inspected))
                    inspected = 0;
                inspected++;
                InspectionQty = inspected; // This updates the backing field and raises OnPropertyChanged
                txtInspectionQty.Text = inspected.ToString();

                // Increment OkCount or NgCount depending on inspection status
                bool allOk = resultsWithStatus.All(r => r.Value.IsOk);
                if (allOk)
                {
                    int okCount;
                    if (!int.TryParse(txtOkCount.Text, out okCount))
                    okCount = 0;
                    okCount++;
                    OkCount = okCount;
                    txtOkCount.Text = okCount.ToString();
                }
                else
                {
                    int ngCount;
                    if (!int.TryParse(txtNgCount.Text, out ngCount))
                        ngCount = 0;
                    ngCount++;
                    NgCount = ngCount;
                    txtNgCount.Text = ngCount.ToString();
                }

                currentMasterCount++;


                // After incrementing counts, call master expiration check
                CheckMasterExpirationDuringMeasurement();
            });

           
        }



        private void LoadMasterInspectionProgressBars(string partNumber)
        {
            try
            {
                // 1️⃣ Get merged data from SQL
                var masterReadings = dataStorageService.GetMasterReadingByPart(partNumber);

                if (masterReadings == null || masterReadings.Count == 0)
                {
                    MessageBox.Show("No master inspection parameters found for this part.",
                                    "Master Inspection", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                // 2️⃣ Convert to PartReadingDataModel (so LoadProgressBars can use it)
                parameterData = masterReadings.Select(m => new PartReadingDataModel
                {
                    Para_No = m.Para_No,
                    Parameter = m.Parameter,
                    Nominal = m.Nominal,
                    RTolPlus = m.RTolPlus,
                    RTolMinus = m.RTolMinus,
                    D_Name=m.D_Name
                    
                }).ToList();

                // 3️⃣ Call your existing progress bar loader
                LoadProgressBars();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading master inspection progress bars: {ex.Message}",
                                "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }




        private void AutoManual_Loaded(object sender, RoutedEventArgs e)
        {
            // Load saved mode from DB and set toggle button state
            int modeBit = dataStorageService.GetAutoManualBit();
            AutoManualToggle.IsChecked = (modeBit == 1);
        }

        private async void ResultPage_Loaded(object sender, RoutedEventArgs e)
        {
            this.Focusable = true;
            this.Focus();

            this.PreviewKeyDown += ResultPage_PreviewKeyDown;
            InitializeValveDataAndUI();
            InitializeDataGrid();
            NotifyStatus("Initialization....");
            //try
            //{
            //    // Ensure PLC and Probe Connection asynchronously when page loads
            //    bool connected = await _masterService.EnsureConnectionAsync();
            //    if (connected)
            //    {
            //        NotifyStatus("PLC and Probe Connected Successfully");
            //    }
            //    else
            //    {
            //        MessageBox.Show("Failed to connect to PLC and Probe", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            //    }
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show($"Error during connection: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            //}
        }

        // Helper method to add timeout to PLC connection




        private void ResultPage_Unloaded(object sender, RoutedEventArgs e)
        {
            _masterService.Dispose();
        }


        private void InitializeValveDataAndUI()
        {
            try
            {
                var activeParts = dataStorageService.GetActiveParts();
                if (activeParts == null || activeParts.Count == 0)
                {
                    MessageBox.Show("No active parts found.");
                    return;
                }

                activePartNumber = activeParts[0].Para_No ?? string.Empty;
                parameterData = dataStorageService.GetPartConfigByPartNumber(activePartNumber);

                if (parameterData == null || parameterData.Count == 0)
                {
                    MessageBox.Show($"No parameters found for active part {activePartNumber}");
                    return;
                }

                fullToShortMap = parameterData.ToDictionary(p => p.Parameter, p => p.ShortName ?? p.Parameter);
                LoadProgressBars();
                SwitchProgressBarBtn.Content = useFirstDesign ? "Switch to Design 2" : "Switch to Design 1";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading data: {ex.Message}");
            }
        }

        private void ValveReadingsGrid_AutoGeneratingColumn(object sender, DataGridAutoGeneratingColumnEventArgs e)
        {
            e.Column.Width = new DataGridLength(1, DataGridLengthUnitType.Star);

            var headerStyle = new Style(typeof(DataGridColumnHeader));
            headerStyle.Setters.Add(new Setter(HorizontalContentAlignmentProperty, HorizontalAlignment.Center));
            e.Column.HeaderStyle = headerStyle;
        }



        private void InitializeDataGrid()
        {
            _measurementDataTable = new DataTable();

            // Add serial number column
            _measurementDataTable.Columns.Add("No", typeof(int));

            // Assuming parameterData contains the list from the DB with parameter and ShortName info
            foreach (var param in parameterData)
            {
                string columnName = param.ShortName;  // use ShortName
                _measurementDataTable.Columns.Add(columnName, typeof(string));
            }


            ValveReadingsGrid.AddHandler(DataGridColumnHeader.ClickEvent,
                new RoutedEventHandler(DataGridColumnHeader_Click));

            // Start with empty DataGrid (no rows)
            ValveReadingsGrid.ItemsSource = _measurementDataTable.DefaultView;
        }


        private void DataGridColumnHeader_Click(object sender, RoutedEventArgs e)
        {
            if (e.OriginalSource is DataGridColumnHeader header && header.Column != null)
            {
                string parameterName = header.Column.Header.ToString();
                if (parameterName == "No") return;

                var values = GetLast10NumericValues(parameterName);
                //if (values.Count == 0)
                //{
                //    MessageBox.Show($"No readings available for {parameterName}.",
                //        "No Data", MessageBoxButton.OK, MessageBoxImage.Information);
                //    return;
                //}

                // Show big custom window
                var win = new VariationWindow(parameterName, values)
                {
                    Owner = Window.GetWindow(ValveReadingsGrid)
                };
                win.ShowDialog();
            }
        }




        private List<double> GetLast10NumericValues(string parameterName)
        {
            if (_measurementDataTable == null || !_measurementDataTable.Columns.Contains(parameterName))
                return new List<double>();

            return _measurementDataTable.AsEnumerable()
                .Take(10)
                .Select(r =>
                {
                    double val;
                    return double.TryParse(r[parameterName]?.ToString(), out val) ? val : double.NaN;
                })
                .Where(v => !double.IsNaN(v))
                .ToList();
        }




        private int _globalSerialCounter = 1; // Initialize in your class (not in method)

        private void LoadDataGrid(Dictionary<string, double> latestValues, Dictionary<string, ParameterResult> resultsWithStatus)
        {
            if (_measurementDataTable == null || latestValues == null) return;

            Dispatcher.Invoke(() =>
            {
                var row = _measurementDataTable.NewRow();
                row["No"] = _globalSerialCounter++;

                // Assign data using full-to-short name map to ensure values map to correct short-named columns
                foreach (var kvp in latestValues) // kvp.Key = full parameter name
                {
                    if (fullToShortMap.TryGetValue(kvp.Key, out string shortName) &&
                        _measurementDataTable.Columns.Contains(shortName))
                    {
                        row[shortName] = kvp.Value.ToString("F3");
                    }
                }

                // Fill any columns not assigned with empty string
                foreach (DataColumn col in _measurementDataTable.Columns)
                {
                    if (col.ColumnName == "No") continue;
                    if (row[col.ColumnName] == DBNull.Value)
                    {
                        row[col.ColumnName] = "";
                    }
                }

                _measurementDataTable.Rows.InsertAt(row, 0);

                while (_measurementDataTable.Rows.Count > MaxRows) // MaxRows = 10 in your code
                    _measurementDataTable.Rows.RemoveAt(_measurementDataTable.Rows.Count - 1);

                // No need to reassign ItemsSource repeatedly - just update layout
                ValveReadingsGrid.UpdateLayout();

                // Highlight cells with NG status in red
                ValveReadingsGrid.Dispatcher.BeginInvoke(new Action(() =>
                {
                    int rowIndex = 0; // Only color newest row

                    foreach (var kvp in fullToShortMap)
                    {
                        string shortHeader = kvp.Value;
                        var column = ValveReadingsGrid.Columns.FirstOrDefault(c => c.Header.ToString() == shortHeader);
                        if (column == null) continue;

                        var cellContent = column.GetCellContent(ValveReadingsGrid.Items[rowIndex]);
                        if (cellContent == null) continue;

                        var cell = FindParent<DataGridCell>(cellContent);
                        if (cell == null) continue;

                        // Default color (reset first)
                        cell.Foreground = Brushes.Black;

                        // Now apply red if NG
                        if (resultsWithStatus.TryGetValue(kvp.Key, out var result) && !result.IsOk)
                        {
                            cell.Foreground = Brushes.Red;
                        }
                    }
                }), System.Windows.Threading.DispatcherPriority.Background);

            });
        }


        private T FindParent<T>(DependencyObject child) where T : DependencyObject
        {
            DependencyObject parentObject = VisualTreeHelper.GetParent(child);
            if (parentObject == null) return null;
            if (parentObject is T parent) return parent;
            return FindParent<T>(parentObject);
        }









        // Dictionary to hold references to dynamically created progress bar controls keyed by parameter name

        /// <summary>
        /// Dynamically load progress bars into ProgressBarContainer based on parameterData.
        /// Tracks controls in _progressBarControls dictionary for later value updates.
        /// </summary>
        private void LoadProgressBars()
        {
            ProgressBarContainer.Children.Clear();
            _progressBarControls.Clear();

            // ⚠️ Set columns based on design selection
            if (useFirstDesign)
            {
                ProgressBarContainer.Columns = 9;   // DESIGN 1
            }
            else
            {
                ProgressBarContainer.Columns = 11;    // DESIGN 2 (set your value)
            }

            foreach (var param in parameterData)
            {
                UserControl progressBar;

                double min = param.Nominal - param.RTolMinus;
                double max = param.Nominal + param.RTolPlus;
                double mean = param.Nominal;

                if (useFirstDesign)
                {
                    // ---- DESIGN 1 ----
                    var pb = new ResultProgressBar();
                    pb.Margin = new Thickness(5);
                    pb.Width = 200;
                    pb.Height = 160;

                    pb.ParameterName = param.D_Name;
                    pb.MinValue = min;
                    pb.MaxValue = max;
                    pb.MeanValue = mean;
                    pb.Value = 0;

                    progressBar = pb;
                }
                else
                {
                    // ---- DESIGN 2 ----
                    var pb = new ProgresBarControl();
                    pb.Margin = new Thickness(5);
                    pb.Width = 200;
                    pb.Height = 160;

                    pb.Min = min;
                    pb.Mean = mean;
                    pb.Max = max;
                    pb.Value = 0;
                    pb.Title = param.Parameter;

                    progressBar = pb;
                }

                ProgressBarContainer.Children.Add(progressBar);
                _progressBarControls[param.Parameter] = progressBar;
            }
        }




        private void UpdateProgressBarsWithStatus(Dictionary<string, ParameterResult> resultsWithStatus)
        {
            foreach (var kvp in resultsWithStatus)
            {
                if (_progressBarControls.TryGetValue(kvp.Key, out var control))
                {
                    if (control is ResultProgressBar pb)
                    {
                        pb.UpdateValue(kvp.Value.Value, kvp.Value.IsOk);
                    }
                    else if (control is ProgresBarControl pb2)
                    {
                        pb2.Value = kvp.Value.Value;
                        // Optional: add IsOk property and color logic to ProgresBarControl if desired
                    }
                }
            }
        }






        private void ToggleBtn_Click(object sender, RoutedEventArgs e)
        {
            if (_showLeft)
            {
                LeftColumn.Width = new GridLength(0);
                RightColumn.Width = new GridLength(1, GridUnitType.Star);
            }
            else
            {
                LeftColumn.Width = new GridLength(1, GridUnitType.Star);
                RightColumn.Width = new GridLength(0);
            }
            _showLeft = !_showLeft;
        }

        private void SwitchProgressBar_Click(object sender, RoutedEventArgs e)
        {
            useFirstDesign = !useFirstDesign;
            LoadProgressBars();
            SwitchProgressBarBtn.Content = useFirstDesign ? "Switch to Design 2" : "Switch to Design 1";
        }


        // Mastering toggle
        private ToggleButton? _activeToggleButton = null;

        // Tracks the currently active toggle

        // Enable all toggles
        private void EnableAllToggles()
        {
            MasterToggle.IsEnabled = true;
            MasterInspectionToggleButton.IsEnabled = true;
            MeasurementToggle.IsEnabled = true;

            // Re-enable Auto/Manual toggle only when no operation is active
            AutoManualToggle.IsEnabled = true;
        }

        // Disable all toggles except the active one
        private void DisableOtherToggles(ToggleButton active)
        {
            MasterToggle.IsEnabled = (active == MasterToggle);
            MasterInspectionToggleButton.IsEnabled = (active == MasterInspectionToggleButton);
            MeasurementToggle.IsEnabled = (active == MeasurementToggle);

            // Disable Auto/Manual toggle whenever an operation toggle is active
            AutoManualToggle.IsEnabled = false;
        }

        // ========================= MASTER TOGGLE =========================
        private async void MasterToggle_Checked(object sender, RoutedEventArgs e)
        {
            if (!(sender is ToggleButton toggleButton)) return;

            if (_activeToggleButton != null && _activeToggleButton != toggleButton)
            {
                MessageBox.Show("Please turn off the other active operation before starting this.",
                                "Operation Active", MessageBoxButton.OK, MessageBoxImage.Warning);
                toggleButton.IsChecked = false;
                return;
            }

            StartBitMatchCheck(); // Start monitoring Auto/Manual bit after starting measurement
            _activeToggleButton = toggleButton;
            DisableOtherToggles(toggleButton);

            try
            {
                ResetAllResult();
                _masterService.IsMasteringStage = true;
                _masterService._continueMeasurement = false;
                ResetMeasurementFieldsAndProgressBars();

                _currentMode = ProcedureMode.Mastering;

                await _masterService.MasterCheckProcedureAsync(ProcedureMode.Mastering);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error starting mastering: {ex.Message}");
            }
            finally
            {
                toggleButton.IsChecked = false;
                _masterService._continueMeasurement = false;
                _activeToggleButton = null;
                EnableAllToggles();
            }
        }

        // ========================= MASTER INSPECTION TOGGLE =========================
        private async void MasterInspectionToggleButton_Checked(object sender, RoutedEventArgs e)
        {
            if (!(sender is ToggleButton toggleButton)) return;

            if (!_masterService.IsConnected)
            {
                MessageBox.Show("PLC not connected. Cannot perform inspection.");
                toggleButton.IsChecked = false;
                return;
            }

            if (_activeToggleButton != null && _activeToggleButton != toggleButton)
            {
                MessageBox.Show("Please turn off the other active operation before starting this.",
                                "Operation Active", MessageBoxButton.OK, MessageBoxImage.Warning);
                toggleButton.IsChecked = false;
                return;
            }

            _activeToggleButton = toggleButton;
            DisableOtherToggles(toggleButton);

            try
            {
                LoadMasterInspectionProgressBars(activePartNumber);


                ResetAllResult();
                _masterService.IsMasteringStage = false;
                _masterService._continueMeasurement = false;
                ResetMeasurementFieldsAndProgressBars();
                _currentMode = ProcedureMode.MasterInspection;

                StartBitMatchCheck();

                await _masterService.MasterCheckProcedureAsync(ProcedureMode.MasterInspection);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error during master inspection: {ex.Message}");
            }
            finally
            {
                toggleButton.IsChecked = false;
                _masterService._continueMeasurement = false;
                _activeToggleButton = null;
                EnableAllToggles();
            }
        }


        // ========================= MEASUREMENT TOGGLE =========================
        private async void MeasurementToggle_Checked(object sender, RoutedEventArgs e)
        {
            if (!(sender is ToggleButton toggleButton)) return;

            if (!_masterService.IsConnected)
            {
                MessageBox.Show("PLC not connected. Cannot perform measurement.");
                toggleButton.IsChecked = false;
                return;
            }

            if (_activeToggleButton != null && _activeToggleButton != toggleButton)
            {
                MessageBox.Show("Please turn off the other active operation before starting this.",
                                "Operation Active", MessageBoxButton.OK, MessageBoxImage.Warning);
                toggleButton.IsChecked = false;
                return;
            }

            // Removed expiration check here per requirement
            StartBitMatchCheck(); // Start monitoring Auto/Manual bit after starting measurement
            currentMasterCount = 0;

            _activeToggleButton = toggleButton;
            DisableOtherToggles(toggleButton);


            try
            {
                InitializeValveDataAndUI();
                _masterService.SetPlcDevice("M101", 0);
                _masterService.SetPlcDevice("M102", 0); // General rejection

                _masterService.IsMasteringStage = false;
                _masterService._continueMeasurement = true;
                ResetMeasurementFieldsAndProgressBars();

                Dispatcher.Invoke(() =>
                {
                    _measurementDataTable?.Clear();
                    _globalSerialCounter = 1;
                    
                });

                _currentMode = ProcedureMode.Measurement;

                await _masterService.MasterCheckProcedureAsync(ProcedureMode.Measurement);
                await _masterService.RunMeasurementCycleAsync();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error during measurement: {ex.Message}");
            }
            finally
            {
                toggleButton.IsChecked = false;
                _masterService._continueMeasurement = false;
                _activeToggleButton = null;
                EnableAllToggles();
            }
        }



        // ========================= TOGGLE UNCHECKED HANDLER =========================
        private void ToggleButton_Unchecked(object sender, RoutedEventArgs e)
        {
            if (!(sender is ToggleButton toggleButton)) return;

            // Stop measurement if Measurement toggle turned off
            if (toggleButton == MeasurementToggle)
            {
                ResetAllResult();
                _masterService.SetPlcDevice("M301", 0); // _masterService.SetPlcDevice("M10", 0); // 

                _masterService._continueMeasurement = false;

                UpdateInspectionDataAsync();

                //ResetAllPlcBits();
            }

    bitMatchCheckTimer?.Stop();

            try
            {
                _masterService.SetPlcDevice("M300", 0);
                NotifyStatus(".");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error turning off PLC devices: {ex.Message}", "PLC Error",
                                MessageBoxButton.OK, MessageBoxImage.Error);
            }

            if (toggleButton == _activeToggleButton)
            {
                _activeToggleButton = null;
            }

            // Re-enable all toggles including Auto/Manual
            EnableAllToggles();
        }


        private void AutoManualToggle_Checked(object sender, RoutedEventArgs e)
        {
            // Save Auto = 1 to DB when toggled on
            dataStorageService.UpdateAutoManualBit(1);
        }

        private void AutoManualToggle_Unchecked(object sender, RoutedEventArgs e)
        {
            // Save Manual = 0 to DB when toggled off
            dataStorageService.UpdateAutoManualBit(0);
        }


        private void ProbeReadingHandler(object? sender, ProbeReadingEventArgs e)
        {
            Application.Current.Dispatcher.Invoke(() =>
            {
                MessageBox.Show($"Probe {e.ModuleId}: {e.Value:0.000}");
            });
        }




        private int _inspectionQty;
        public int InspectionQty
        {
            get => _inspectionQty;
            set { _inspectionQty = value; OnPropertyChanged(nameof(InspectionQty)); }
        }

        private int _okCount;
        public int OkCount
        {
            get => _okCount;
            set { _okCount = value; OnPropertyChanged(nameof(OkCount)); }
        }

        private int _ngCount;
        public int NgCount
        {
            get => _ngCount;
            set { _ngCount = value; OnPropertyChanged(nameof(NgCount)); }
        }

        //public void AddPartInspectionResults(Dictionary<string, ParameterResult> parameterResults)
        //{
        //    InspectionQty++; // increment total parts inspected

        //    bool partIsOk = parameterResults.All(r => r.Value.IsOk);
        //    if (partIsOk)
        //        OkCount++;
        //    else
        //        NgCount++;
        //}


        private void UpdateMeasurementResult(TextBox textBox, double value, bool isOk)
        {
            if (textBox == null) return;

            textBox.Text = value.ToString("0.000");

            // Change background color, keep text foreground default
            textBox.Background = isOk ? Brushes.LimeGreen : Brushes.IndianRed;
        }


        private void UpdateMeasurementLabel(TextBlock label, bool isOk)
        {
            if (label == null) return;

        }

        private void UpdateFieldIfExists(Dictionary<string, ParameterResult> results, string key, TextBox textBox, TextBlock label)
        {
            if (results.TryGetValue(key, out var result))
            {
                UpdateMeasurementResult(textBox, result.Value, result.IsOk);
                UpdateMeasurementLabel(label, result.IsOk);
            }
        }


        private void UpdateMeasurementFields(Dictionary<string, ParameterResult> resultsWithStatus)
        {
            UpdateFieldIfExists(resultsWithStatus, "Groove Diameter", GrooveDiaBox, GrooveDiaLabel);
            UpdateFieldIfExists(resultsWithStatus, "Groove Position", GroovePo, GroovePoLabel);
            UpdateFieldIfExists(resultsWithStatus, "Stem Dia Near Groove", STDG, STDGLabel);
            UpdateFieldIfExists(resultsWithStatus, "Stem Dia Near Undercut", StemDiaBox, StemDiaLabel);
            UpdateFieldIfExists(resultsWithStatus, "Head Diameter", HeadDiaBox, HeadDiaLabel);
            UpdateFieldIfExists(resultsWithStatus, "Face Runout", FaceRunout, FaceRunoutLabel);
            UpdateFieldIfExists(resultsWithStatus, "Seat Height", SeatHeightBox, SeatHeightLabel);
            UpdateFieldIfExists(resultsWithStatus, "Seat Runout", SeatRunout, SeatRunoutLabel);
            UpdateFieldIfExists(resultsWithStatus, "Datum to End", DatumToEndBox, DatumToEndLabel);
            UpdateFieldIfExists(resultsWithStatus, "Datum to Groove", DatumToGrooveBox, DatumToGrooveLabel);
            UpdateFieldIfExists(resultsWithStatus, "Stem Taper", StemTaper, StemTaperLabel);
            UpdateFieldIfExists(resultsWithStatus, "Straightness", Straightness, StraightnessLabel);
            UpdateFieldIfExists(resultsWithStatus, "Overall Length", OverallLengthBox, OverallLengthLabel);
            UpdateFieldIfExists(resultsWithStatus, "End Face Runout", HeadRunout, HeadRunoutLabel);
            UpdateFieldIfExists(resultsWithStatus, "Ovality Head", HeadOver, HeadOverLabel);
        }

        private void ResetMeasurementFieldsAndProgressBars()
        {
            // Reset all measurement TextBox values and backgrounds
            Brush defaultBg = Brushes.White;

            string zeroText = "00";

            GrooveDiaBox.Text = zeroText; GrooveDiaBox.Background = defaultBg;
            GroovePo.Text = zeroText; GroovePo.Background = defaultBg;
            STDG.Text = zeroText; STDG.Background = defaultBg;
            StemDiaBox.Text = zeroText; StemDiaBox.Background = defaultBg;
            HeadDiaBox.Text = zeroText; HeadDiaBox.Background = defaultBg;
            FaceRunout.Text = zeroText; FaceRunout.Background = defaultBg;
            SeatHeightBox.Text = zeroText; SeatHeightBox.Background = defaultBg;
            SeatRunout.Text = zeroText; SeatRunout.Background = defaultBg;
            DatumToEndBox.Text = zeroText; DatumToEndBox.Background = defaultBg;
            DatumToGrooveBox.Text = zeroText; DatumToGrooveBox.Background = defaultBg;
            StemTaper.Text = zeroText; StemTaper.Background = defaultBg;
            Straightness.Text = zeroText; Straightness.Background = defaultBg;
            OverallLengthBox.Text = zeroText; OverallLengthBox.Background = defaultBg;
            HeadRunout.Text = zeroText; HeadRunout.Background = defaultBg;
            HeadOver.Text = zeroText; HeadOver.Background = defaultBg;

            // Reset progress bar controls for both designs if present
            foreach (var control in _progressBarControls.Values)
            {
                if (control is ResultProgressBar pb1)
                {
                    pb1.Value = 0;       // reset value
                    pb1.UpdateValue(0, null); // clear any color status
                }
                else if (control is ProgresBarControl pb2)
                {
                    pb2.Value = 0;       // reset value
                                         // Optionally add a method to clear IsOk/Color state in ProgresBarControl if implemented
                }
            }
        }



        private void ResultPage_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                // MessageBox.Show("ESC pressed in Result Page", "Key Pressed", MessageBoxButton.OK, MessageBoxImage.Information);
                _masterService._continueMeasurement = false; // Stop measurement
                e.Handled = true;
                ResetAllPlcBits();
                _masterService.Dispose();
                // Additional ESC handling logic here
                HandleEscKeyAction();
                ResetAllResult();

                NotifyStatus(".");

            }
        }



        private void HandleEscKeyAction()
        {
            Window currentWindow = Window.GetWindow(this);
            if (currentWindow != null)
            {
                // Assuming your window has a container named MainContentGrid
                var mainContentGrid = currentWindow.FindName("MainContentGrid") as Grid;
                if (mainContentGrid != null)
                {
                    mainContentGrid.Children.Clear();
                    var resultPage = new Dashboard();

                    resultPage.HorizontalAlignment = HorizontalAlignment.Stretch;
                    resultPage.VerticalAlignment = VerticalAlignment.Stretch;
                    NotifyStatus(".");
                    mainContentGrid.Children.Add(resultPage);
                }
            }
        }

        private void ResetAllResult()
        {
            _masterService.SetPlcDevice("M301", 0); // General rejection
            _masterService.SetPlcDevice("M302", 0); // SRO rejection
            _masterService.SetPlcDevice("M303", 0); // STDIA rejection
            _masterService.SetPlcDevice("M304", 0); // Seat Height rejection
            _masterService.SetPlcDevice("M305", 0); // Groove Diameter/Position rejection
            _masterService.SetPlcDevice("M306", 0); // Groove Diameter/Position rejection

        }


        private void ResetAllPlcBits()
        {
            string[] bitsToReset =
            {
                    "M400", "M100", "M300", "M10", "M14", "M301"
                };

            try
            {
                foreach (var bit in bitsToReset)
                {
                    _masterService.SetPlcDevice(bit, 0);
                }
            }
            catch (Exception ex)
            {
                // Show a single message for all errors
                MessageBox.Show(
                    $"PLC Reset Failed.\nError: {ex.Message}",
                    "PLC Error",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error
                );

                // Close current window/page
                Application.Current.Dispatcher.Invoke(() =>
                {
                    Window currentWindow = Window.GetWindow(this);
                    if (currentWindow != null)
                    {
                        currentWindow.Close();
                    }
                });
            }
        }




        private void StartShiftTimer()
        {
            shiftTimer = new DispatcherTimer();
            shiftTimer.Interval = TimeSpan.FromSeconds(1);  // checks every minute
            shiftTimer.Tick += ShiftTimer_Tick;
            shiftTimer.Start();

            UpdateShiftDisplay(); // Initial update
        }

        private void ShiftTimer_Tick(object sender, EventArgs e)
        {
            UpdateShiftDisplay();
        }

        private void UpdateShiftDisplay()
        {
            try
            {
                // 🧠 Get Auto/Manual bit from data storage
                var autoList = dataStorageService.GetActiveBit();
                var autoControl = autoList.FirstOrDefault(c =>
                    string.Equals(c.Code, "L1", StringComparison.OrdinalIgnoreCase));

                // ✅ If Auto mode is active (bit == 1), generate lot number automatically
                if (autoControl != null && Convert.ToInt32(autoControl.Bit) == 1)
                {
                    string currentShift = GetShiftCode();
                    string dateShift = DateTime.Now.ToString("ddMMyy") + currentShift;  // Example: 031125A
                    txtLotNo.Text = dateShift;
                }
                else
                {
                    // ❌ Manual mode: keep user-entered value, do NOT overwrite
                    txtLotNo.Text = txtLotNo.Text;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error updating shift display: {ex.Message}",
                    "Auto/Manual Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private string GetShiftCode()
        {
            TimeSpan current = DateTime.Now.TimeOfDay;

            TimeSpan shiftAStart = new TimeSpan(6, 0, 0);
            TimeSpan shiftAEnd = new TimeSpan(13, 59, 59);
            TimeSpan shiftBStart = new TimeSpan(14, 0, 0);
            TimeSpan shiftBEnd = new TimeSpan(21, 59, 59);
            TimeSpan shiftCStart = new TimeSpan(22, 0, 0);
            TimeSpan shiftCEnd = new TimeSpan(5, 59, 59);

            if (current >= shiftAStart && current <= shiftAEnd)
                return "A";
            else if (current >= shiftBStart && current <= shiftBEnd)
                return "B";
            else
                return "C";
        }



        private DispatcherTimer expirationTimer1;

        /// <summary>
        /// Checks master expiration: stops on count, starts (and stops) timer for time.
        /// Call at the start of measurement and after each cycle.
        /// </summary>
        // Automatically turns off mastering when inspection count reaches expiration limit
        // Master expiration using only Count mode
        // Call this method after each inspection count update (e.g., after each measurement)
        private int currentMasterCount = 0;

        public void CheckMasterExpirationDuringMeasurement()
        {
            var (mode, masterCount, _) = dataStorageService.GetMasterExpiration();

            // Only check for count mode expiration
            if (mode == 1 )
            {
                if (currentMasterCount >= masterCount)
                {
                    // Auto-turn off measurement
                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        if (MasterToggle.IsChecked == true)
                            MasterToggle.IsChecked = false;
                    });

                    _masterService._continueMeasurement = false;

                    MessageBox.Show(
                        "Master count limit reached. Measurement has been automatically stopped.",
                        "Master Expiration",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning
                    );

                    // Reset the counter for next time
                    currentMasterCount = 0;
                }
            }
        }



        private DispatcherTimer bitMatchCheckTimer;
        private int? lastSoftwareBitValue = null;
        private int? lastPlcBitValue = null;
        private bool wasPreviousMismatch = false;

        private void StartBitMatchCheck()
        {
            bitMatchCheckTimer = new DispatcherTimer();
            bitMatchCheckTimer.Interval = TimeSpan.FromSeconds(1);
            bitMatchCheckTimer.Tick += BitMatchCheckTimer_Tick;
            bitMatchCheckTimer.Start();
        }

        private void BitMatchCheckTimer_Tick(object sender, EventArgs e)
        {
            var autoList = dataStorageService.GetActiveBit();
            var autoControl = autoList.FirstOrDefault(c => string.Equals(c.Description, "Auto/Manual", StringComparison.OrdinalIgnoreCase));

            if (autoControl == null)
                return; // Can't check without software bit

            int softwareBit = autoControl.Bit; // Software bit value
            int plcBit = _masterService.GetPlcDeviceBit("X14"); // PLC bit

            if (softwareBit == plcBit)
            {
                // Bits match - reset mismatch state
                wasPreviousMismatch = false;
            }
            else
            {
                // Bits do not match
                if (wasPreviousMismatch)
                {
                    // Consecutive mismatch detected - show error and stop measurement
                    bitMatchCheckTimer.Stop(); // Stop checking further

                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        MessageBox.Show("Auto/Manual bit mismatch detected consecutively. PLC is in Manual mode or out of sync.",
                                        "Error", MessageBoxButton.OK, MessageBoxImage.Error);

                        if (_masterService._continueMeasurement)
                        {
                            _masterService._continueMeasurement = false;
                            if (MeasurementToggle.IsChecked == true)
                                MeasurementToggle.IsChecked = false;
                        }
                    });
                }
                else
                {
                    // First mismatch, set flag and wait next tick to confirm
                    wasPreviousMismatch = true;
                }
            }
        }



        private async void HandleDeleteLatestMeasurementRowAndCounts()
        {
            // 1. Get the latest row's status before removal
            if (_measurementDataTable == null || _measurementDataTable.Rows.Count == 0) return;

            var latestRow = _measurementDataTable.Rows[0]; // Top row is latest
            string status = latestRow["No"]?.ToString();

            // 2. Delete row from DB using DataStorageService
            bool deleted = await dataStorageService.DeleteLatestMeasurementReadingAsync(_model, _lotNo, int.Parse(_userId));
            if (!deleted)
            {
                MessageBox.Show("No record deleted. No matching data found.", "Delete", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            // 3. Remove row from DataTable/UI and refresh_measurementDataTable
            _measurementDataTable.Rows.RemoveAt(0);
            ValveReadingsGrid.Items.Refresh();

            // 4. Decrement counts from UI and DB
            int qty = 0, okCount = 0, ngCount = 0;
            int.TryParse(txtInspectionQty.Text, out qty);
            int.TryParse(txtOkCount.Text, out okCount);
            int.TryParse(txtNgCount.Text, out ngCount);

            if (qty > 0) qty--;
            if (status == "OK" && okCount > 0) okCount--;
            else if (status == "NG" && ngCount > 0) ngCount--;

            txtInspectionQty.Text = qty.ToString();
            txtOkCount.Text = okCount.ToString();
            txtNgCount.Text = ngCount.ToString();

            // Call DataStorageService to update counts in database
            await dataStorageService.UpdateInspectionCountsAsync(_model, _lotNo, _userId, qty, okCount);

            // 5. Optionally decrement global serial counter
            if (_globalSerialCounter > 1) _globalSerialCounter--;

            ValveReadingsGrid.UpdateLayout();
        }


    }
}

