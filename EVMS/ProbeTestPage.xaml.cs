using Microsoft.Data.SqlClient;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Configuration;
using System.Globalization;
using System.IO.Ports;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace EVMS
{
    public partial class ProbeTestPage : UserControl, INotifyPropertyChanged
    {
        private readonly string _connectionString;
        private SerialPort? _serial;
        private CancellationTokenSource? _serialCts;
        private Task? _serialReadTask;

        // Data collections
        public ObservableCollection<string> PartNumbers { get; } = new();
        public ObservableCollection<TestProbeRow> Probes { get; } = new();

        private string _selectedPartNo = "";
        public string SelectedPartNo
        {
            get => _selectedPartNo;
            set
            {
                if (_selectedPartNo != value)
                {
                    _selectedPartNo = value;
                    OnPropertyChanged();
                }
            }
        }

        public ProbeTestPage()
        {
            InitializeComponent();
            DataContext = this;
            _connectionString = ConfigurationManager.ConnectionStrings["EVMSDb"]?.ConnectionString ?? "";

            Loaded += ProbeTestPage_Loaded;
        }

        //==============================================================
        //  🔥 Auto Start Reading After Page Loads
        //==============================================================
        private async void ProbeTestPage_Loaded(object? sender, RoutedEventArgs e)
        {
            await LoadPartNumbersAsync();

            if (PartNumbers.Count == 0)
            {
                StatusText.Text = "⚠ No part found in DB!";
                return;
            }

            SelectedPartNo = PartNumbers[0];
            PartNumbersCombo.SelectedIndex = 0;

            StatusText.Text = "📥 Loading probe configuration...";
            await LoadProbeDataFromInstallationDataAsync(SelectedPartNo);

            ConnectToCom3();
            StatusText.Text = "🔄 Starting Live Read (40 samples each)...";

            StartLiveReadingAuto();   // <<< Auto Test Start
        }



        private async Task LoadPartNumbersAsync()
        {
            try
            {
                if (string.IsNullOrEmpty(_connectionString))
                {
                    MessageBox.Show("Connection string 'EVMSDb' not found in config.");
                    return;
                }

                using var con = new SqlConnection(_connectionString);
                await con.OpenAsync();

                const string sql = @"
            SELECT Para_No 
            FROM PART_ENTRY 
            WHERE ActivePart = 1
            ORDER BY Para_No";

                using var cmd = new SqlCommand(sql, con);

                PartNumbers.Clear();

                using var reader = await cmd.ExecuteReaderAsync();
                while (await reader.ReadAsync())
                {
                    // NO GetString(0) here
                    object raw = reader.GetValue(0);                    // System.Int32
                    string paraNo = Convert.ToString(raw, CultureInfo.InvariantCulture) ?? "";

                    if (!string.IsNullOrWhiteSpace(paraNo))
                        PartNumbers.Add(paraNo);
                }

                if (PartNumbers.Count > 0)
                    SelectedPartNo = PartNumbers[0];
            }
            catch (Exception ex)
            {
                MessageBox.Show("DB Error " + ex.Message);
            }
        }


        private async Task LoadProbeDataFromInstallationDataAsync(string partNo)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(partNo)) return;

                using var con = new SqlConnection(_connectionString);
                await con.OpenAsync();

                string q = "SELECT ParameterName,BoxId,ChannelId FROM ProbeInstallationData WHERE PartNo=@p ORDER BY ParameterName, ChannelId";
                using var cmd = new SqlCommand(q, con);
                cmd.Parameters.AddWithValue("@p", partNo);

                using var rd = await cmd.ExecuteReaderAsync();
                Probes.Clear();

                while (await rd.ReadAsync())
                {
                    var name = rd["ParameterName"]?.ToString() ?? "";
                    var box = Convert.ToInt32(rd["BoxId"]);
                    var ch = Convert.ToInt32(rd["ChannelId"]);

                    Probes.Add(new TestProbeRow
                    {
                        Name = name,
                        Title = $"{name} (CH {ch})",
                        BoxId = box,
                        Channel = ch,
                        LiveValue = 0,
                        Status = "Ready"
                    });
                }

                StatusText.Text = $"🧪 Loaded {Probes.Count} Probes";
                UpdateSummary();
            }
            catch (Exception ex) { MessageBox.Show("Load Error " + ex.Message); }
        }

        //==============================================================
        //  SERIAL + AUTO 40 READINGS COLLECT
        //==============================================================
        private void StartLiveReadingAuto()
        {
            if (_serial == null || !_serial.IsOpen)
            {
                StatusText.Text = "❌ Serial Not Connected";
                return;
            }

            // reset existing readings
            foreach (var p in Probes) p.Readings.Clear();

            _serialCts?.Cancel();
            _serialCts = new CancellationTokenSource();
            _serialReadTask = Task.Run(() => CollectFixedReadingsAsync(_serialCts.Token));
        }

        public List<double> Readings { get; set; } = new();

        private async Task CollectFixedReadingsAsync(CancellationToken token)
        {
            const int TARGET = 40; // required reading count per probe

            try
            {
                while (!token.IsCancellationRequested)
                {
                    var boxes = Probes.Select(p => p.BoxId).Distinct().Where(b => b > 0).ToList();

                    foreach (var box in boxes)
                    {
                        token.ThrowIfCancellationRequested();
                        await ReadBoxOnce(box, token);

                        foreach (var p in Probes.Where(x => x.BoxId == box))
                        {
                            double value = p.LiveValue;

                            // Validate actual reading (allow negatives)
                            bool valid = !double.IsNaN(value);

                            // Store only if:
                            // 1. Value valid (not NaN)
                            // 2. Count < target
                            // 3. Value is new or different from last saved
                            if (valid && p.Readings.Count < TARGET)
                            {
                                if (p.Readings.Count == 0 || p.Readings.Last() != value)
                                {
                                    p.Readings.Add(value);
                                }
                            }
                        }

                    }

                    // Update UI summary in main thread
                    Application.Current.Dispatcher.Invoke(UpdateSummary);

                    // stop only when every probe reaches exact TARGET
                    if (Probes.All(p => p.Readings.Count == TARGET))
                        break;

                    await Task.Delay(60, token); // sampling delay
                }

                Application.Current.Dispatcher.Invoke(() =>
                {
                    StatusText.Text = "🎉 Completed — Live Reading Stopped (All readings reached 40)";
                });
            }
            catch (Exception ex)
            {
                Application.Current.Dispatcher.Invoke(() => StatusText.Text = $"❌ Error: {ex.Message}");
            }
            finally
            {
                StopSerial(); // stop serial only when test completed
            }
        }


        //==============================================================
        //  SERIAL FUNCTIONS
        //==============================================================
        private void ConnectToCom3()
        {
            try
            {
                try
                {
                    _serial?.Close();
                    _serial?.Dispose();
                }
                catch { /* ignore */ }

                _serial = new SerialPort("COM3", 115200, Parity.None, 8, StopBits.One)
                {
                    ReadTimeout = 500,
                    WriteTimeout = 500,
                    NewLine = "\r"
                };
                _serial.Open();
            }
            catch (Exception ex)
            {
                StatusText.Text = $"⚠ COM3 Open Fail: {ex.Message}";
            }
        }

        private async Task ReadBoxOnce(int box, CancellationToken token)
        {
            if (_serial == null || !_serial.IsOpen) return;

            string cmd = $"*{box:D3}VALL#\r";
            try
            {
                _serial.DiscardInBuffer();
                _serial.DiscardOutBuffer();
                _serial.Write(cmd);
            }
            catch
            {
                MarkProbesBoxError(box);
                return;
            }

            string resp = await Task.Run(() => ReadFullResponse(200, token));
            if (string.IsNullOrWhiteSpace(resp))
            {
                MarkProbesBoxError(box);
                return;
            }

            double[] values = ParseResponse(resp);
            UpdateProbesFromBox(box, values);
        }

        private void MarkProbesBoxError(int box)
        {
            foreach (var p in Probes.Where(p => p.BoxId == box))
                UpdateProbeUI(p, 0, "ERR", false);
        }

        private void UpdateProbesFromBox(int box, double[] vals)
        {
            foreach (var p in Probes.Where(x => x.BoxId == box))
            {
                if (p.Channel < 1 || p.Channel > 4)
                {
                    UpdateProbeUI(p, double.NaN, "CH ERR", false);
                    continue;
                }

                double value = vals.Length > p.Channel ? vals[p.Channel] : double.NaN;
                if (double.IsNaN(value))
                {
                    UpdateProbeUI(p, double.NaN, "ERR", false);
                    continue;
                }

                bool ok = value >= 0.2 && value <= 1.2;
                UpdateProbeUI(p, value, $"{value:0.000}", ok);
            }
        }

        private void UpdateProbeUI(TestProbeRow p, double val, string status, bool ok)
        {
            Application.Current.Dispatcher.Invoke(() =>
            {
                p.LiveValue = val;
                p.Status = status;
                p.InRange = ok;
            });
        }

        private string ReadFullResponse(int timeoutMs = 80, CancellationToken token = default)
        {
            if (_serial == null) return "";
            string resp = "";
            var sw = System.Diagnostics.Stopwatch.StartNew();

            try
            {
                while (sw.ElapsedMilliseconds < timeoutMs)
                {
                    token.ThrowIfCancellationRequested();
                    if (_serial.BytesToRead > 0)
                    {
                        resp += _serial.ReadExisting();
                        if (resp.Contains("#")) break;
                    }
                    Thread.Sleep(1);
                }
            }
            catch (OperationCanceledException) { return ""; }
            catch { }
            return resp;
        }

        private double[] ParseResponse(string r)
        {
            double[] ch = new double[5];
            for (int i = 1; i <= 4; i++) ch[i] = double.NaN;

            if (string.IsNullOrWhiteSpace(r)) return ch;

            int idx = r.IndexOf("VALL", StringComparison.OrdinalIgnoreCase);
            if (idx >= 0) r = r[(idx + 4)..];

            r = r.Replace("#", "").Replace("\r", "").Replace("\n", "").Trim();

            var matches = Regex.Matches(r, @"C\d{2}([-+]?\d*\.?\d+)");
            for (int i = 0; i < matches.Count && i < 4; i++)
            {
                string valueStr = matches[i].Groups[1].Value;
                if (double.TryParse(valueStr, NumberStyles.Any, CultureInfo.InvariantCulture, out double v))
                    ch[i + 1] = v;
            }
            return ch;
        }

        private void StopSerial()
        {
            try
            {
                _serialCts?.Cancel();
                _serialReadTask = null;
                _serialCts?.Dispose();
                _serialCts = null;

                if (_serial != null)
                {
                    try { if (_serial.IsOpen) _serial.Close(); }
                    catch { }
                    try { _serial.Dispose(); }
                    catch { }
                    _serial = null;
                }

                // optionally clear live values (keeps readings)
                Application.Current.Dispatcher.Invoke(() =>
                {
                    foreach (var p in Probes)
                    {
                        p.InRange = false;
                        p.Status = ""; // or keep last
                    }
                    UpdateSummary();
                });
            }
            catch { }
        }

        //==============================================================
        private void UpdateSummary()
        {
            TotalProbesText.Text = Probes.Count.ToString();
            TotalReadingsText.Text = Probes.Sum(p => p.Readings.Count).ToString();
        }

        public event PropertyChangedEventHandler? PropertyChanged;
        private void OnPropertyChanged([CallerMemberName] string? n = null) =>
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(n));
    }

    //==============================================================
    // MODEL
    //==============================================================
    public class TestProbeRow : INotifyPropertyChanged
    {
        private string _title = "";
        private string _name = "";
        private int _boxId;
        private int _channel;
        private double _liveValue;
        private string _status = "";
        private bool _inRange;
        private readonly List<double> _readings = new();

        public string Title { get => _title; set { _title = value; OnPropertyChanged(); } }
        public string Name { get => _name; set { _name = value; OnPropertyChanged(); } }
        public int BoxId { get => _boxId; set { _boxId = value; OnPropertyChanged(); } }
        public int Channel { get => _channel; set { _channel = value; OnPropertyChanged(); } }
        public double LiveValue { get => _liveValue; set { _liveValue = value; OnPropertyChanged(); } }
        public string Status { get => _status; set { _status = value; OnPropertyChanged(); } }
        public bool InRange { get => _inRange; set { _inRange = value; OnPropertyChanged(); } }

        // Use List<double> internally but expose count and min/max computed properties
        public List<double> Readings => _readings;
        public int ReadingsCount => _readings.Count;
        public double MinValue => _readings.Any() ? _readings.Min() : 0;
        public double MaxValue => _readings.Any() ? _readings.Max() : 0;

        public event PropertyChangedEventHandler? PropertyChanged;
        protected virtual void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    // ---------------- CONVERTER INCLUDED HERE (TOP-LEVEL)
    public class BoolToColorConverter : System.Windows.Data.IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
            => (value is bool b && b) ? Brushes.LightGreen : Brushes.IndianRed;

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
            => throw new NotImplementedException();
    }
}
