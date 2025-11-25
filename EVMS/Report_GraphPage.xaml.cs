using EVMS.Service;
using PdfSharpCore.Drawing;
using ScottPlot;
using ScottPlot.WPF;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using MathNet.Numerics.Statistics;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Linq;
using PdfSharpCore.Pdf;
using System.Windows.Media.Imaging;

namespace EVMS
{
    public partial class Report_GraphPage : UserControl, INotifyPropertyChanged
    {
        private readonly DataStorageService _dataService;

        // Observable collections bound to UI
        public ObservableCollection<string> ActiveParts { get; set; } = new();
        public ObservableCollection<string> LotNumbers { get; set; } = new();
        public ObservableCollection<string> ParametersOptions { get; set; } = new();
        public ObservableCollection<string> Operators { get; set; } = new();
        public ObservableCollection<string> DesignOptions { get; set; } = new();

        // Selected properties
        private string _selectedPartNo;
        private string _selectedLotNo;
        private string _selectedParameter;
        private string _selectedOperator;
        private string _selectedDesign;

        private DateTime? _selectedDateTimeFrom = DateTime.Now.AddDays(-7);
        private DateTime? _selectedDateTimeTo = DateTime.Now;
        private bool _isPrinting = false; // flag to prevent multiple dialogs


        public event PropertyChangedEventHandler PropertyChanged;

        // Mapping UI parameter name -> model property name (exact)
        private static readonly Dictionary<string, string> ParameterToColumn = new()
        {
            ["Overall Length"] = "OL",
            ["Datum to End"] = "DE",
            ["Head Diameter"] = "HD",
            ["Groove Position"] = "GP",
            ["Stem Dia Near Groove"] = "STDG",
            ["Stem Dia Near Undercut"] = "STDU",
            ["Groove Diameter"] = "GIR_DIA",
            ["Straightness"] = "STN",
            ["Ovality SDG"] = "Ovality_SDG",
            ["Ovality SDU"] = "Ovality_SDU",
            ["Ovality Head"] = "Ovality_Head",
            ["Stem Taper"] = "Stem_Taper",
            ["End Face Runout"] = "EFRO",
            ["Face Runout"] = "Face_Runout",
            ["Seat Height"] = "SH"
        };

        public Report_GraphPage()
        {
            InitializeComponent();

            _dataService = new DataStorageService();

            LoadDesignOptions();
            LoadActiveParts();

            Loaded += Report_GraphPage_Loaded;

            DataContext = this;

            // Attach keydown event
            this.PreviewKeyDown += Report_GraphPage_PreviewKeyDown;
            this.Focusable = true;
            this.Focus();
        }

        #region Properties (bindings)
        public string SelectedPartNo
        {
            get => _selectedPartNo;
            set
            {
                if (_selectedPartNo != value)
                {
                    _selectedPartNo = value;
                    OnPropertyChanged(nameof(SelectedPartNo));
                    _ = OnSelectedPartNoChangedAsync();
                }
            }
        }

        public string SelectedLotNo
        {
            get => _selectedLotNo;
            set { _selectedLotNo = value; OnPropertyChanged(nameof(SelectedLotNo)); }
        }

        public string SelectedParameter
        {
            get => _selectedParameter;
            set { _selectedParameter = value; OnPropertyChanged(nameof(SelectedParameter)); }
        }

        public string SelectedOperator
        {
            get => _selectedOperator;
            set { _selectedOperator = value; OnPropertyChanged(nameof(SelectedOperator)); }
        }

        public string SelectedDesign
        {
            get => _selectedDesign;
            set
            {
                if (_selectedDesign != value)
                {
                    _selectedDesign = value;
                    OnPropertyChanged(nameof(SelectedDesign));
                    UpdateChartVisibility();
                }
            }
        }

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
        #endregion

        #region Initialization & loaders
        private void Report_GraphPage_Loaded(object sender, RoutedEventArgs e)
        {
            if (ScottPlotControl != null)
                ScottPlotControl.Visibility = Visibility.Collapsed;
        }

        private void LoadDesignOptions()
        {
            DesignOptions.Clear();

            // Existing
            DesignOptions.Add("Line Chart");
            DesignOptions.Add("Histogram");

            // NEW GRAPH TYPES
            DesignOptions.Add("Normal Distribution");             // Histogram with Normal Curve
            DesignOptions.Add("Run Chart");     // I-Chart / X-Chart
            DesignOptions.Add("Gage R&R");             // Trend / Run Chart

            // Default selected
            SelectedDesign = DesignOptions.FirstOrDefault();
        }

        private void LoadActiveParts()
        {
            try
            {
                var parts = _dataService.GetActiveParts();
                ActiveParts.Clear();
                ActiveParts.Add("All");
                if (parts != null)
                {
                    foreach (var p in parts)
                    {
                        if (!string.IsNullOrWhiteSpace(p.Para_No))
                            ActiveParts.Add(p.Para_No);
                    }
                }
                SelectedPartNo = "All";
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[LoadActiveParts] {ex.Message}");
            }
        }


        private void Report_GraphPage_PreviewKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            // Check if 'P' is pressed
            if (e.Key == System.Windows.Input.Key.P)
            {
                // Generate PDF
                GeneratePlotPdf();

                // Prevent further processing
                e.Handled = true;
                return;
            }
        }


        private void GeneratePlotPdf()
        {
            if (ScottPlotControl == null)
                return;

            try
            {
                // 1. Render ScottPlot to bitmap
                Bitmap bmp = ScottPlotControl.Plot.Render();

                // 2. Ask user for file save location
                Microsoft.Win32.SaveFileDialog dlg = new Microsoft.Win32.SaveFileDialog();
                dlg.Filter = "PDF Files (*.pdf)|*.pdf";
                dlg.FileName = $"{SelectedParameter}_{DateTime.Now:yyyyMMdd_HHmm}.pdf";

                if (dlg.ShowDialog() != true) return;

                string filename = dlg.FileName;

                // 3. Create PDF document
                PdfDocument doc = new PdfDocument();
                doc.Info.Title = "ScottPlot Graph";

                PdfPage page = doc.AddPage();
                page.Width = XUnit.FromPoint(bmp.Width);
                page.Height = XUnit.FromPoint(bmp.Height);

                using (XGraphics gfx = XGraphics.FromPdfPage(page))
                using (MemoryStream ms = new MemoryStream())
                {
                    bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                    ms.Seek(0, SeekOrigin.Begin);

                    // PdfSharpCore expects a delegate that returns a Stream
                    XImage img = XImage.FromStream(() => new MemoryStream(ms.ToArray()));

                    gfx.DrawImage(img, 0, 0, page.Width, page.Height);
                }

                // 4. Save PDF
                using (FileStream fs = new FileStream(filename, FileMode.Create, FileAccess.Write))
                {
                    doc.Save(fs);
                }

                MessageBox.Show($"PDF saved to: {filename}", "Print", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error generating PDF");
            }
        }


        private async Task OnSelectedPartNoChangedAsync()
        {
            if (string.IsNullOrEmpty(SelectedPartNo)) return;
            await LoadLotNumbersAsync(SelectedPartNo);
            LoadParameters(SelectedPartNo);
            await LoadOperatorsAsync(SelectedPartNo);
        }

        private async Task ReloadLotAndOperatorAsync()
        {
            await LoadLotNumbersAsync(SelectedPartNo);
            await LoadOperatorsAsync(SelectedPartNo);
        }

        public async Task LoadLotNumbersAsync(string partNo)
        {
            try
            {
                var lots = await _dataService.GetLotNumbersByPartAndDateRangeAsync(
                    partNo == "All" ? null : partNo, SelectedDateTimeFrom, SelectedDateTimeTo);

                LotNumbers.Clear();
                LotNumbers.Add("All");
                if (lots != null)
                {
                    foreach (var lot in lots)
                    {
                        if (!string.IsNullOrWhiteSpace(lot) && !LotNumbers.Contains(lot))
                            LotNumbers.Add(lot);
                    }
                }
                SelectedLotNo = "All";
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[LoadLotNumbersAsync] {ex.Message}");
            }
        }

        public async Task LoadOperatorsAsync(string partNo)
        {
            try
            {
                var ops = await _data_service_safe_getops(partNo, SelectedDateTimeFrom, SelectedDateTimeTo);
                Operators.Clear();
                Operators.Add("All");
                if (ops != null)
                {
                    foreach (var op in ops)
                    {
                        if (!string.IsNullOrWhiteSpace(op) && !Operators.Contains(op))
                            Operators.Add(op);
                    }
                }
                SelectedOperator = "All";
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[LoadOperatorsAsync] {ex.Message}");
            }
        }

        private async Task<IEnumerable<string>> _data_service_safe_getops(string partFilter, DateTime? from, DateTime? to)
        {
            try
            {
                return await _dataService.GetOperatorsByPartAndDateRangeAsync(partFilter, from, to) ?? Enumerable.Empty<string>();
            }
            catch { return Enumerable.Empty<string>(); }
        }

        private void LoadParameters(string part)
        {
            try
            {
                ParametersOptions.Clear();
                var config = _dataService.GetPartConfig(part == "All" ? null : part);
                if (config != null)
                {
                    foreach (var c in config)
                        ParametersOptions.Add(c.Parameter);
                }
                SelectedParameter = ParametersOptions.FirstOrDefault();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[LoadParameters] {ex.Message}");
            }
        }
        #endregion

        #region Charting

        private async void OnSubmitClicked(object sender, RoutedEventArgs e)
        {
            await LoadGraphAsync();
        }

        // --------------------------------------------------------------------
        // NOTE:
        // - Your original implementation is preserved below as LoadGraphAsync_Old.
        // - The new active method is LoadGraphAsync which implements the unified
        //   X = measurement, Y = count logic for Line, Bar, Histogram, while
        //   preserving Gage R&R behavior.
        // --------------------------------------------------------------------






        private void PlotGageRR(List<object> list, string col)
        {
            var plt = ScottPlotControl.Plot;
            plt.Clear();

            var rAndRData = list
                .Select(m => new
                {
                    Part = (string)m.GetType().GetProperty("PartNo")?.GetValue(m),
                    Operator = (string)m.GetType().GetProperty("Operator_ID")?.GetValue(m),
                    Trial = (int?)m.GetType().GetProperty("TrialNo")?.GetValue(m),
                    Measurement = Convert.ToDouble(m.GetType().GetProperty(col)?.GetValue(m) ?? 0)
                }).ToList();

            if (rAndRData.Count == 0)
                return;

            // --- GROUP DATA ---
            var grouped = rAndRData.GroupBy(d => new { d.Operator, d.Part })
                                   .ToDictionary(g => g.Key, g => g.Select(x => x.Measurement).ToList());

            var operators = rAndRData.Select(d => d.Operator).Distinct().OrderBy(o => o).ToList();
            var parts = rAndRData.Select(d => d.Part).Distinct().OrderBy(p => p).ToList();

            // --- R CHART ---
            double[] rValues = new double[operators.Count * parts.Count];
            int idx = 0;
            foreach (var op in operators)
            {
                foreach (var part in parts)
                {
                    var key = new { Operator = op, Part = part };
                    if (grouped.ContainsKey(key))
                    {
                        var measurements = grouped[key];
                        double r = measurements.Max() - measurements.Min();
                        rValues[idx++] = r;
                    }
                    else rValues[idx++] = 0;
                }
            }

            var rBar = plt.AddBar(rValues);
            rBar.FillColor = System.Drawing.Color.Blue;
            rBar.BarWidth = 0.5;
            plt.Title("R Chart by Operator");
            plt.XTicks(Enumerable.Range(1, rValues.Length).Select(i => (double)i).ToArray());

            // --- XBAR CHART ---
            double[] xbarValues = new double[operators.Count * parts.Count];
            idx = 0;
            foreach (var op in operators)
            {
                foreach (var part in parts)
                {
                    var key = new { Operator = op, Part = part };
                    if (grouped.ContainsKey(key))
                    {
                        var measurements = grouped[key];
                        xbarValues[idx++] = measurements.Average();
                    }
                    else xbarValues[idx++] = 0;
                }
            }

            var xbarBar = plt.AddBar(xbarValues);
            xbarBar.FillColor = System.Drawing.Color.Green;
            xbarBar.BarWidth = 0.5;
            plt.Title("Xbar Chart by Operator");

            ScottPlotControl.Refresh();
        }


        // --------------------------
        // NEW: Unified LoadGraphAsync
        // --------------------------
        private async Task LoadGraphAsync()
    {
    try
        {
            if (string.IsNullOrEmpty(SelectedParameter) || string.IsNullOrEmpty(SelectedPartNo))
                return;

            string part = SelectedPartNo == "All" ? null : SelectedPartNo;
        string lot = SelectedLotNo == "All" ? null : SelectedLotNo;
        string oper = SelectedOperator == "All" ? null : SelectedOperator;

        var config = _dataService.GetPartConfig(part)?.FirstOrDefault(c => c.Parameter == SelectedParameter);
        if (config == null) return;

        var list = await _data_service_safe_getops(part, SelectedDateTimeFrom, SelectedDateTimeTo) == null
            ? await _dataService.GetMeasurementReadingsAsync(part, lot, oper, SelectedDateTimeFrom, SelectedDateTimeTo)
            : await _dataService.GetMeasurementReadingsAsync(part, lot, oper, SelectedDateTimeFrom, SelectedDateTimeTo);
        // (Above: kept your existing call; you may replace with direct call if needed)
        if (list == null) return;

        string col = ParameterToColumn.ContainsKey(SelectedParameter) ? ParameterToColumn[SelectedParameter] : null;
        if (col == null) return;

        // Collect numeric values
        List<double> valuesList = new();
        foreach (var m in list)
        {
            var p = m.GetType().GetProperty(col);
            if (p != null && double.TryParse(p.GetValue(m)?.ToString(), out double val))
                valuesList.Add(val);
        }

        if (valuesList.Count == 0) return;

        // Tolerances
        double nominal = config.Nominal;
        double LSL = config.Nominal - config.RTolMinus;
        double USL = config.Nominal + config.RTolPlus;

        var plt = ScottPlotControl.Plot;
        plt.Clear();

        // -------------------------
        // Prepare histogram bins
        // -------------------------
        double dataMin = valuesList.Min();
        double dataMax = valuesList.Max();

        // Determine min/max to cover both data and tolerance band
        double minX = Math.Min(LSL, dataMin);
        double maxX = Math.Max(USL, dataMax);

        // Protect against degenerate range
        if (Math.Abs(maxX - minX) < 1e-12)
        {
            minX -= 1.0;
            maxX += 1.0;
        }

        // Bin count heuristic
        int binCount = 20;
        binCount = Math.Min(binCount, Math.Max(5, valuesList.Count)); // at least 5 bins, no more than sample count

        double binSize = (maxX - minX) / binCount;

        double[] binCenters = new double[binCount];
        double[] counts = new double[binCount];
        for (int i = 0; i < binCount; i++)
            binCenters[i] = minX + (i * binSize) + binSize / 2.0;

        // Tally counts into manual bins
        foreach (double v in valuesList)
        {
            int binIndex = (int)((v - minX) / binSize);
            if (binIndex < 0) binIndex = 0;
            if (binIndex >= binCount) binIndex = binCount - 1;
            counts[binIndex]++;
        }

        // -------------------------
        // Compute statistics (full formulas)
        // -------------------------
        double mean = valuesList.Average();

        // Population sigma (divide by N) — matches screenshot style
        double sigma = 0;
        if (valuesList.Count > 0)
            sigma = Math.Sqrt(valuesList.Sum(v => Math.Pow(v - mean, 2)) / valuesList.Count);

        double sixSigma = 6.0 * sigma;

        double minValue = dataMin;
        double maxValue = dataMax;

        double classInterval = binSize;

        // Cp: (USL - LSL) / (6 * sigma) — handle sigma==0
        double cp = double.NaN;
        if (sigma > 0)
            cp = (USL - LSL) / (6.0 * sigma);

        // Cpk:
        double cpu = double.NaN;
        double cpl = double.NaN;
        double cpk = double.NaN;
        if (sigma > 0)
        {
            cpu = (USL - mean) / (3.0 * sigma);
            cpl = (mean - LSL) / (3.0 * sigma);
            cpk = Math.Min(cpu, cpl);
        }

        // Pp / Ppk (use sample sigma: divide by N-1)
        double sigma_p = double.NaN;
        if (valuesList.Count > 1)
            sigma_p = Math.Sqrt(valuesList.Sum(v => Math.Pow(v - mean, 2)) / (valuesList.Count - 1));

        double pp = double.NaN;
        if (!double.IsNaN(sigma_p) && sigma_p > 0)
            pp = (USL - LSL) / (6.0 * sigma_p);

        double ppu = double.NaN, ppl = double.NaN, ppk = double.NaN;
        if (!double.IsNaN(sigma_p) && sigma_p > 0)
        {
            ppu = (USL - mean) / (3.0 * sigma_p);
            ppl = (mean - LSL) / (3.0 * sigma_p);
            ppk = Math.Min(ppu, ppl);
        }

        // Format values safely
        string fmtD(double d) => double.IsNaN(d) ? "-" : d.ToString("0.###");

        // -------------------------
        // Draw charts based on SelectedDesign
        // -------------------------

        // LINE CHART: counts vs bin centers (distribution line)
        if (SelectedDesign == "Line Chart")
        {
            var scatter = plt.AddScatter(binCenters, counts, lineWidth: 2);
            scatter.Color = System.Drawing.Color.Blue;

            plt.Title($"{SelectedParameter} - Distribution (Line)");
            plt.XLabel("Measurement Value");
            plt.YLabel("Count");
        }
        // BAR CHART: bar histogram
        else if (SelectedDesign == "Histogram")
        {
            var bar = plt.AddBar(counts, binCenters);
            bar.BarWidth = binSize * 0.9;
            bar.FillColor = System.Drawing.Color.SteelBlue;

            plt.Title($"{SelectedParameter} - Histogram (Bar Chart)");
            plt.XLabel("Measurement Value");
            plt.YLabel("Count");
        }
                else if (SelectedDesign == "Normal Distribution")
                {
                    var bar = plt.AddBar(counts, binCenters);
                    bar.BarWidth = binSize * 0.9;
                    bar.FillColor = System.Drawing.Color.FromArgb(128, System.Drawing.Color.LightGreen);
                    bar.BorderLineWidth = 1;
                    bar.BorderColor = System.Drawing.Color.DarkGreen;

                    plt.Title($"{SelectedParameter} - Histogram + Normal Distribution");
                    plt.XLabel("Measurement Value");
                    plt.YLabel("Count");

                    // ---- Normal Distribution Curve ----
                    if (sigma > 0 && valuesList.Count > 1)
                    {
                        int numPoints = 600;
                        double[] x = new double[numPoints];
                        double[] y = new double[numPoints];

                        double step = (maxX - minX) / (numPoints - 1);

                        for (int i = 0; i < numPoints; i++)
                        {
                            x[i] = minX + i * step;
                            y[i] = Math.Exp(-0.5 * Math.Pow((x[i] - mean) / sigma, 2))
                                   / (sigma * Math.Sqrt(2 * Math.PI));
                        }

                        // scale curve to histogram height
                        double maxHist = counts.Max();
                        double maxCurve = y.Max();
                        for (int i = 0; i < y.Length; i++)
                            y[i] = y[i] / maxCurve * maxHist;

                        plt.AddScatter(x, y, lineWidth: 3,
                            color: System.Drawing.Color.Black)
                           .Label = "Normal Curve";
                    }

                    
                    

                    plt.Legend(true);

                    // adjust view
                    plt.SetAxisLimits(minX, maxX, 0, counts.Max() * 1.20);
                }

                // Gage R&R: keep your original implementation (copied in)
                else if (SelectedDesign == "Gage R&R")
                {
                    PlotGageRR(list.Cast<object>().ToList(), col);
                    return;   // ⬅ Important: stop histogram drawing
                }
                else if (SelectedDesign == "Run Chart")
                {
                    var run = ScottPlotControl.Plot;
                    run.Clear();

                    // X-axis sample numbers
                    double[] xs = Enumerable.Range(1, valuesList.Count)
                                            .Select(i => (double)i)
                                            .ToArray();
                    double[] ys = valuesList.ToArray();

                    // Plot line + markers
                    var series = run.AddScatter(xs, ys,
                        color: System.Drawing.Color.DarkBlue,
                        lineWidth: 2, markerSize: 6);
                    series.Label = SelectedParameter;

                    // Mean line
                    run.AddHorizontalLine(mean, System.Drawing.Color.Green, 2)
                       .Label = $"Mean {mean:0.###}";

                    // Spec limits
                    run.AddHorizontalLine(USL, System.Drawing.Color.Red, 2)
                       .Label = $"USL {USL:0.###}";
                    run.AddHorizontalLine(LSL, System.Drawing.Color.Red, 2)
                       .Label = $"LSL {LSL:0.###}";

                    run.Title($"{SelectedParameter} - Run Chart");
                    run.XLabel("Sample Number");
                    run.YLabel("Measurement Value");

                    run.Legend(true);

                    // Auto Y scaling with small padding
                    run.SetAxisLimits(
                        xMin: 1,
                        xMax: valuesList.Count,
                        yMin: Math.Min(valuesList.Min(), LSL) - 0.05,
                        yMax: Math.Max(valuesList.Max(), USL) + 0.05
                    );

                    ScottPlotControl.Refresh();
                    return;   // ✅ stop other chart drawing
                }



                // -------------------------
                // Draw vertical tolerance lines
                // -------------------------
                try
                {
            var lslLine = plt.AddVerticalLine(LSL, System.Drawing.Color.Red, 2);
            lslLine.Label = $"LSL {LSL:0.###}";
            var uslLine = plt.AddVerticalLine(USL, System.Drawing.Color.Red, 2);
            uslLine.Label = $"USL {USL:0.###}";
            var nominalLine = plt.AddVerticalLine(nominal, System.Drawing.Color.Green, 2);
            nominalLine.Label = $"Nominal {nominal:0.###}";

            plt.Legend(true);
        }
        catch
        {
            // ignore legend errors
        }

                // -------------------------
                // Compose statistics text
                // -------------------------
                string statsText =
      $"USL: {fmtD(USL)}    LSL: {fmtD(LSL)}    Tol: {fmtD(USL - LSL)}\n" +
      $"Min: {fmtD(minValue)}    Max: {fmtD(maxValue)}\n" +
      $"Sigma: {fmtD(sigma)}    6 Sigma: {fmtD(sixSigma)}\n" +
      $"Cp: {fmtD(cp)}    Cpk: {fmtD(cpk)}\n" +
      $"Pp: {fmtD(pp)}    Ppk: {fmtD(ppk)}";

                // Compute height above histogram
                double tallest = counts.Length > 0 ? counts.Max() : 0;
                double yTop = tallest * 1.18;  // 18% above highest bar

                // RIGHT SIDE position
                double xRight = maxX - (maxX - minX) * 0.02; // 2% left from right border

                // Add text aligned at TOP-RIGHT
                var txt = plt.AddText(statsText, xRight, yTop);
                txt.Alignment = ScottPlot.Alignment.UpperRight;   // <— FIXED!
                txt.FontSize = 14;
                txt.FontBold = true;
                txt.Color = System.Drawing.Color.Black;

                // -------------------------
                // Final axis limits and refresh
                // -------------------------
                plt.SetAxisLimits(
            xMin: Math.Min(LSL, minX) - binSize * 1.5,
            xMax: Math.Max(USL, maxX) + binSize * 1.5,
            yMin: 0
        );

        ScottPlotControl.Refresh();
    }
    catch (Exception ex)
    {
        MessageBox.Show(ex.Message, "Graph Error");
    }
}



        // ---------------------------------------------------------
        // ORIGINAL METHOD: preserved (renamed) so nothing is removed
        // ---------------------------------------------------------
        private async Task LoadGraphAsync_Old()
        {
            try
            {
                if (string.IsNullOrEmpty(SelectedParameter) || string.IsNullOrEmpty(SelectedPartNo))
                    return;

                string part = SelectedPartNo == "All" ? null : SelectedPartNo;
                string lot = SelectedLotNo == "All" ? null : SelectedLotNo;
                string oper = SelectedOperator == "All" ? null : SelectedOperator;

                var config = _dataService.GetPartConfig(part)?.FirstOrDefault(c => c.Parameter == SelectedParameter);
                if (config == null) return;

                var list = await _dataService.GetMeasurementReadingsAsync(part, lot, oper, SelectedDateTimeFrom, SelectedDateTimeTo);
                if (list == null) return;

                string col = ParameterToColumn[SelectedParameter];

                List<double> y = new();
                foreach (var m in list)
                {
                    var p = m.GetType().GetProperty(col);
                    if (p != null && double.TryParse(p.GetValue(m)?.ToString(), out double val))
                        y.Add(val);
                }

                if (y.Count == 0) return;

                // --------------------
                // Reset Plot
                // --------------------
                var plt = ScottPlotControl.Plot;
                plt.Clear();

                double[] values = y.ToArray();
                double[] xs = Enumerable.Range(1, values.Length).Select(i => (double)i).ToArray();

                // --------------------
                // Tolerances
                // --------------------
                double nominal = config.Nominal;
                double LSL = config.Nominal - config.RTolMinus;
                double USL = config.Nominal + config.RTolPlus;

                // =========================================================
                //  LINE CHART
                // =========================================================
                if (SelectedDesign == "Line Chart")
                {
                    plt.AddScatter(xs, values, color: System.Drawing.Color.Blue, lineWidth: 2, markerSize: 5);

                    plt.AddHorizontalLine(nominal, System.Drawing.Color.Green, 2).Label = $"Nominal {nominal}";
                    plt.AddHorizontalLine(LSL, System.Drawing.Color.Red, 2).Label = $"LSL {LSL}";
                    plt.AddHorizontalLine(USL, System.Drawing.Color.Red, 2).Label = $"USL {USL}";

                    plt.Title($"{SelectedParameter} - Line Chart");
                    plt.XLabel("Record No");
                    plt.YLabel("Value");
                    plt.Legend();
                }

                // =========================================================
                //  BAR CHART
                // =========================================================
                else if (SelectedDesign == "Bar Chart")
                {
                    for (int i = 0; i < values.Length; i++)
                    {
                        double xCenter = xs[i];
                        double barWidth = 0.8;

                        bool ok = values[i] >= LSL && values[i] <= USL;

                        var color = ok ? System.Drawing.Color.Green : System.Drawing.Color.Red;

                        // Manual rectangle bar
                        var bar = plt.AddBar(
                            values: new double[] { values[i] },
                            positions: new double[] { xCenter }
                        );

                        bar.FillColor = color;
                    }

                    // Tolerances
                    plt.AddHorizontalLine(nominal, System.Drawing.Color.Green, 2).Label = $"Nominal {nominal}";
                    plt.AddHorizontalLine(LSL, System.Drawing.Color.Red, 2).Label = $"LSL {LSL}";
                    plt.AddHorizontalLine(USL, System.Drawing.Color.Red, 2).Label = $"USL {USL}";

                    plt.Title($"{SelectedParameter} - Bar Chart");
                    plt.XLabel("Record No");
                    plt.YLabel("Value");
                    plt.Legend();
                }

                // =========================================================
                //  CAPABILITY PLOT (Cp, Cpk Distribution) - v4.1.67 compatible
                // =========================================================

                else if (SelectedDesign == "Gage R&R")
                {
                    // Step 1: Prepare dataset
                    var rAndRData = list
                        .Select(m => new
                        {
                            Part = m.PartNo,
                            Operator = m.Operator_ID,
                            Trial = m.TrialNo,
                            Measurement = Convert.ToDouble(m.GetType().GetProperty(col)?.GetValue(m) ?? 0)
                        }).ToList();

                    if (rAndRData.Count == 0)
                        return;

                    var parts = rAndRData.Select(d => d.Part).Distinct().ToList();
                    var operators = rAndRData.Select(d => d.Operator).Distinct().ToList();

                    // Step 2: Calculate averages
                    double grandMean = rAndRData.Average(d => d.Measurement);

                    // Part means
                    var partMeans = rAndRData.GroupBy(d => d.Part)
                                              .ToDictionary(g => g.Key, g => g.Average(d => d.Measurement));

                    // Operator means
                    var opMeans = rAndRData.GroupBy(d => d.Operator)
                                            .ToDictionary(g => g.Key, g => g.Average(d => d.Measurement));

                    // Repeatability (EV)
                    double ev = 0;
                    foreach (var g in rAndRData.GroupBy(d => new { d.Part, d.Operator }))
                    {
                        var measurements = g.Select(d => d.Measurement).ToList();
                        if (measurements.Count > 1)
                        {
                            double mean = measurements.Average();
                            double variance = measurements.Sum(x => Math.Pow(x - mean, 2)) / measurements.Count;
                            ev += variance;
                        }
                    }
                    ev = Math.Sqrt(ev / (rAndRData.Count > 1 ? rAndRData.Count : 1));

                    // Operator variation (AV)
                    double av = 0;
                    if (opMeans.Count > 1)
                    {
                        double mean = opMeans.Values.Average();
                        av = Math.Sqrt(opMeans.Values.Sum(x => Math.Pow(x - mean, 2)) / opMeans.Count);
                    }

                    // Part-to-Part variation (PV)
                    double pv = 0;
                    if (partMeans.Count > 1)
                    {
                        double mean = partMeans.Values.Average();
                        pv = Math.Sqrt(partMeans.Values.Sum(x => Math.Pow(x - mean, 2)) / partMeans.Count);
                    }

                    // Total variation
                    double tv = Math.Sqrt(ev * ev + av * av + pv * pv);

                    // Step 3: Prepare contributions (%)
                    double[] contributionValues = new double[]
                    {
                            pv / tv * 100,
                            av / tv * 100,
                            ev / tv * 100
                    };

                    for (int i = 0; i < contributionValues.Length; i++)
                        if (double.IsNaN(contributionValues[i])) contributionValues[i] = 0;

                    string[] labels = { "Part-to-Part", "Operator", "Repeatability" };
                    System.Drawing.Color[] colors = { System.Drawing.Color.Orange, System.Drawing.Color.Green, System.Drawing.Color.Blue };

                    // Step 4: Plot each bar manually to assign colors
                    plt.Clear();
                    for (int i = 0; i < contributionValues.Length; i++)
                    {
                        var bar = plt.AddBar(
                            values: new double[] { contributionValues[i] },
                            positions: new double[] { i + 1 }
                        );
                        bar.FillColor = colors[i];
                        bar.BarWidth = 0.6;
                    }

                    plt.XTicks(Enumerable.Range(1, labels.Length).Select(i => (double)i).ToArray(), labels);
                    plt.Title($"Gage R&R %Contribution - {SelectedParameter}");
                    plt.YLabel("% Contribution");
                    plt.SetAxisLimits(yMin: 0, yMax: 100);
                }

                // ------------------------------------------------------
                // ADVANCED Y-AXIS: 4 Points Below LSL, 4 Points Above USL
                // ------------------------------------------------------
                plt.YAxis.LockLimits(false);

                // ------------------------------------------------------
                // ADVANCED Y-AXIS: numeric ticks + tolerance markers + 
                // extra range above and below
                // ------------------------------------------------------

                // Determine span and step
                double span = USL - LSL;
                double step = span / 10.0;   // 10 divisions between LSL & USL

                // Build Y ticks
                List<double> tickPos = new List<double>();
                List<string> tickLbl = new List<string>();

                // ---- 4 points below LSL ----
                for (int i = 4; i >= 1; i--)
                {
                    double v = LSL - i * step;
                    tickPos.Add(v);
                    tickLbl.Add(v.ToString("0.###"));
                }

                // ---- Numeric ticks between LSL → USL ----
                for (int i = 0; i <= 10; i++)
                {
                    double v = LSL + i * step;
                    tickPos.Add(v);

                    // Name special ticks
                    if (Math.Abs(v - LSL) < 0.0001)
                        tickLbl.Add($"LSL ({v:0.###})");
                    else if (Math.Abs(v - USL) < 0.0001)
                        tickLbl.Add($"USL ({v:0.###})");
                    else
                        tickLbl.Add(v.ToString("0.###"));
                }

                // ---- 4 points above USL ----
                for (int i = 1; i <= 4; i++)
                {
                    double v = USL + i * step;
                    tickPos.Add(v);
                    tickLbl.Add(v.ToString("0.###"));
                }

                // Apply tick config
                plt.YAxis.ManualTickPositions(tickPos.ToArray(), tickLbl.ToArray());

                // Set axis limits
                plt.SetAxisLimits(
                    yMin: LSL - 1 * step,
                    yMax: USL + 1 * step
                );

                // Lock Y-axis
                plt.YAxis.LockLimits(true);

                // Lock the Y axis
                plt.YAxis.LockLimits(true);

                // Auto scale
                ScottPlotControl.Refresh();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Graph Error");
            }
        }

        // draws Nominal / LSL / USL lines and legend
        private void AddLimits(Plot plt, dynamic config)
        {
            if (config == null) return;
            try
            {
                double nominal = (double)config.Nominal;
                double lsl = (double)(config.Nominal - config.RTolMinus);
                double usl = (double)(config.Nominal + config.RTolPlus);

                var nominalLine = plt.AddHorizontalLine(nominal, Color.Green, 2);
                nominalLine.Label = $"Nominal ({nominal})";

                var lslLine = plt.AddHorizontalLine(lsl, Color.Red, 2);
                lslLine.Label = $"LSL ({lsl})";

                var uslLine = plt.AddHorizontalLine(usl, Color.Red, 2);
                uslLine.Label = $"USL ({usl})";

                plt.Legend(true);
            }
            catch { }
        }
        #endregion

        #region Misc
        private void Designe_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // simply refresh or update visibility
            UpdateChartVisibility();
        }

        private void UpdateChartVisibility()
        {
            if (ScottPlotControl == null) return;
            ScottPlotControl.Visibility = Visibility.Visible;
        }

        protected void OnPropertyChanged(string name) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        #endregion
    }
}
