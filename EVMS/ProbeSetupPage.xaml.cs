using ActUtlType64Lib;
using Microsoft.Data.SqlClient;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Globalization;
using System.IO.Ports;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;   // ← add this line
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;

namespace EVMS
{
    public partial class ProbeSetupPage : UserControl, INotifyPropertyChanged
    {

        private readonly string _connectionString;
        private readonly DispatcherTimer _timer;
        private readonly DispatcherTimer ioMonitorTimer = new() { Interval = TimeSpan.FromMilliseconds(250) };
        public event Action<string>? StatusMessageChanged;

        // PLC
        private IActUtlType64 plc = new ActUtlType64Class();
        private bool isPlcConnected = false;
        //private bool plcBitState = false;
        //private bool motorRunState = false;

        // IO devices
        private List<IODevice> inputDevices = new();
        private Dictionary<string, Button> inputButtons = new();

        private List<IODevice> outputDevices = new();
        private Dictionary<string, ToggleButton> outputButtons = new();


        // Serial
        private SerialPort? _serial;
        private CancellationTokenSource? _serialCts;
        private readonly int[] _pos = { 4, 16, 28, 40 };

        // Data collections
        public ObservableCollection<string> PartNumbers { get; } = new();
        public ObservableCollection<ProbeRow> Probes { get; } = new();

        private string _selectedPartNo = string.Empty;
        public string SelectedPartNo
        {
            get => _selectedPartNo;
            set
            {
                if (_selectedPartNo != value)
                {
                    _selectedPartNo = value;
                    RaisePropertyChanged();
                    _ = LoadProbeDataFromInstallationDataAsync(_selectedPartNo);
                }
            }
        }

        public ProbeSetupPage()
        {
            InitializeComponent();
            DataContext = this;

            _connectionString = ConfigurationManager.ConnectionStrings["EVMSDb"].ConnectionString!;
            _timer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(100) };
            _timer.Tick += Timer_Tick;
            Loaded += ProbeSetupPage_Loaded;

            this.Loaded += SettingsPage_Loaded;
            this.PreviewKeyDown += SettingsPage_PreviewKeyDown;
        }

        private void SettingsPage_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                HandleEscKeyAction();
                e.Handled = true;
            }
        }

        private void ProbeSetupPage_Loaded(object sender, RoutedEventArgs e)
        {
            ConnectToCom3();
            ConnectPlcOnPageLoad();
        }


        private async void ConnectPlcOnPageLoad()
        {
            NotifyStatus("🔄 Connecting PLC...");

            //plc.ActLogicalStationNumber = 1;
            //var openTask = Task.Run(() => plc.Open());
            //var completed = await Task.WhenAny(openTask, Task.Delay(3000));

            //if (completed == openTask && openTask.Result == 0)
            //{
            //    isPlcConnected = true;
            //    NotifyStatus("✅ PLC Connected");

            //    // ✅ Enable buttons immediately after connect
            //    foreach (var btn in outputButtons.Values)
            //        btn.IsEnabled = true;
            //}
            //else
            //{
            //    isPlcConnected = false;
            //    NotifyStatus("❌ PLC Failed");
            //}
        }




        public List<SerialPortConfigModel> GetSerialPortConfig()
        {
            var list = new List<SerialPortConfigModel>();
            string query = "SELECT ID, ComPort, BaudRate FROM SerialSettings WHERE ID = 1";

            using SqlConnection conn = new(_connectionString);
            using SqlCommand cmd = new(query, conn);

            conn.Open();
            using SqlDataReader reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                list.Add(new SerialPortConfigModel
                {
                    ID = Convert.ToInt32(reader["ID"]),
                    ComPort = reader["ComPort"].ToString(),
                    BaudRate = Convert.ToInt32(reader["BaudRate"])
                });
            }
            return list;
        }



        private void ConnectToCom3()
        {
            try
            {
                _serial?.Close();
                _serial?.Dispose();
                _serial = null;

                // ✅ WHATEVER IN DB - NO COM8 fallback
                var config = GetSerialPortConfig().FirstOrDefault();
                if (config != null && !string.IsNullOrEmpty(config.ComPort))
                {
                    string comPort = config.ComPort;  // DB value

                    _serial = new SerialPort(comPort, 115200, Parity.None, 8, StopBits.One)
                    {
                        ReadTimeout = 500,
                        WriteTimeout = 500
                    };
                    _serial.Open();
                }
            }
            catch { }
        }

        private async void SettingsPage_Loaded(object? sender, RoutedEventArgs e)
        {
            try
            {
                await LoadInputDevicesAsync();
                await LoadOutputDevicesAsync();
                GenerateOutputButtons();
                GenerateInputButtons();

                await LoadPartNumbersAsync();
                if (!string.IsNullOrEmpty(SelectedPartNo))
                    await LoadProbeDataFromInstallationDataAsync(SelectedPartNo);
                // ✅ Wait for PLC connection before initializing
                await Task.Delay(3500); // Give PLC connect time
                await InitializeOutputButtonStatesAsync();

                // ✅ Enable output buttons only if PLC connected
                if (isPlcConnected)
                {
                    foreach (var btn in outputButtons.Values)
                        btn.IsEnabled = true;

                    StartIOMonitoring();
                }



                NotifyStatus(isPlcConnected ? "✅ All systems ready" : "⚠️ PLC not connected");
            }
            catch (Exception ex)
            {
                NotifyStatus($"Error: {ex.Message}");
            }
        }


        #region MOTOR PLC BUTTONS
        private async Task InitializeOutputButtonStatesAsync()
        {
            if (!isPlcConnected) return;

            foreach (var device in outputDevices)
            {
                if (!outputButtons.ContainsKey(device.Bit)) continue;

                try
                {
                    int result = plc.GetDevice(device.Bit, out int value);
                    if (result == 0)
                    {
                        var btn = outputButtons[device.Bit];
                        btn.IsChecked = (value == 1);
                        btn.Background = value == 1 ? Brushes.Green : Brushes.Red;
                    }
                    else
                    {
                        outputButtons[device.Bit].Background = Brushes.Gray;
                    }
                }
                catch
                {
                    if (outputButtons.ContainsKey(device.Bit))
                        outputButtons[device.Bit].Background = Brushes.Gray;
                }
            }
        }



        #endregion

        #region DATABASE LOADING
        private async Task LoadPartNumbersAsync()
        {
            try
            {
                using var con = new SqlConnection(_connectionString);
                await con.OpenAsync();

                // Only active part(s)
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
                    // If Para_No is INT in DB
                    object raw = reader.GetValue(0);
                    string paraNo = Convert.ToString(raw, CultureInfo.InvariantCulture) ?? "";
                    if (!string.IsNullOrWhiteSpace(paraNo))
                        PartNumbers.Add(paraNo);
                }

                if (PartNumbers.Count > 0)
                    SelectedPartNo = PartNumbers[0];
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading active part: {ex.Message}");
            }
        }

        // *** THIS IS NOW THE ONLY PROBE SOURCE ***
        private async Task LoadProbeDataFromInstallationDataAsync(string partNo)
        {
            if (string.IsNullOrEmpty(partNo)) return;

            try
            {
                using var con = new SqlConnection(_connectionString);
                await con.OpenAsync();

                string query = @"
            SELECT ParameterName, BoxId, ChannelId, ProbeName
            FROM ProbeInstallationData
            WHERE PartNo = @P
            ORDER BY 
                CASE 
                    WHEN ProbeName LIKE 'Probe%' THEN 
                        CAST(SUBSTRING(ProbeName, 6, LEN(ProbeName)) AS INT)
                    ELSE 999
                END,
                ProbeName";

                using var cmd = new SqlCommand(query, con);
                cmd.Parameters.AddWithValue("@P", partNo);

                var allRows = new List<(string ParameterName, string ProbeName, int BoxId, int ChannelId)>();
                using var reader = await cmd.ExecuteReaderAsync();

                while (await reader.ReadAsync())
                {
                    allRows.Add((
                        ParameterName: reader.GetString("ParameterName"),
                        ProbeName: reader.GetString("ProbeName"),
                        BoxId: reader.GetInt32("BoxId"),
                        ChannelId: reader.GetInt32("ChannelId")
                    ));
                }

                // Group by ProbeName and sort numerically
                var probeGroups = allRows
                    .GroupBy(x => x.ProbeName)
                    .Select(g => new ProbeRow
                    {
                        ID = g.First().ParameterName,
                        BoxId = g.First().BoxId,
                        Channels = g.Select(x => x.ChannelId).Distinct().OrderBy(c => c).ToList(),
                        Title = g.Key  // Shows exactly "Probe 1", "Probe 2", "Probe 3"
                    })
                    .OrderBy(p => ExtractProbeNumber(p.Title))  // Numeric sort: 1,2,3...
                    .ThenBy(p => p.Title);

                Application.Current.Dispatcher.Invoke(() =>
                {
                    Probes.Clear();
                    foreach (var probe in probeGroups)
                        Probes.Add(probe);
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading probes: {ex.Message}", "Database Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }


        private async Task LoadOutputDevicesAsync()
        {
            outputDevices.Clear();
            using var conn = new SqlConnection(_connectionString);
            await conn.OpenAsync();
            const string sql = "SELECT ID, Description, Bit, IsOutput FROM IODevices WHERE IsOutput = 1";
            using var cmd = new SqlCommand(sql, conn);
            using var reader = await cmd.ExecuteReaderAsync();
            while (await reader.ReadAsync())
            {
                outputDevices.Add(new IODevice
                {
                    Description = reader["Description"]?.ToString() ?? "Unknown",
                    Bit = reader["Bit"]?.ToString() ?? "Unknown",
                    // add IsOutput if you keep that property
                });
            }
        }


        // Helper method for numeric probe sorting
        private int ExtractProbeNumber(string probeName)
        {
            if (probeName.StartsWith("Probe ") && int.TryParse(probeName.Substring(6), out int number))
                return number;
            return 999; // Unknown probes last
        }





        #endregion

        #region SERIAL READING
        private void Timer_Tick(object? sender, EventArgs? e) { }

        private void DisconnectSerialButton_Click(object sender, RoutedEventArgs e)
        {
            StopSerial();
            try
            {
                _serial?.Close();
                _serial?.Dispose();
                _serial = null;
            }
            catch { }
        }

        private void StartSerialButton_Click(object sender, RoutedEventArgs e)
        {
            if (_serial == null || !_serial.IsOpen)
            {
                MessageBox.Show("Connect serial first.");
                return;
            }

            if (_serialCts != null)
            {
                MessageBox.Show("Already running.");
                return;
            }

            _serialCts = new CancellationTokenSource();
            _ = Task.Run(() => SerialLiveLoopAsync(_serialCts.Token));
        }

        private void StopSerial()
        {
            try
            {

                // 1. Cancel immediately
                _serialCts?.Cancel();

                // 2. Wait briefly for graceful exit (300ms max)
                if (_serialCts != null)
                {
                    _serialCts.Token.WaitHandle.WaitOne(300);
                    _serialCts.Dispose();
                }
                _serialCts = null;

            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"StopSerial error: {ex.Message}");
            }
        }


        private async Task SerialLiveLoopAsync(CancellationToken token)
        {
            try
            {
                while (!token.IsCancellationRequested)
                {
                    var boxes = Probes.Select(p => p.BoxId).Distinct().Where(b => b > 0).ToList();

                    foreach (int box in boxes)
                    {
                        // ✅ Check cancellation before each box read
                        token.ThrowIfCancellationRequested();

                        // ✅ PASS TOKEN TO ReadBoxOnce
                        await ReadBoxOnce(box, token);
                        await Task.Delay(10, token);
                    }

                    await Task.Delay(50, token);
                }
            }
            catch (OperationCanceledException)
            {
                System.Diagnostics.Debug.WriteLine("✅ Serial loop cancelled gracefully");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Serial error: {ex.Message}");
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

            // ✅ PASS TOKEN HERE
            string resp = await Task.Run(() => ReadFullResponse(80, token));
            if (string.IsNullOrEmpty(resp))
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
                if (p.Channels == null || p.Channels.Count == 0)
                {
                    UpdateProbeUI(p, 0, "No CH", false);
                    continue;
                }

                int ch = p.Channels[0];  // Use the first (and only) channel

                if (ch < 1 || ch > 4)
                {
                    UpdateProbeUI(p, 0, "CH ERR", false);
                    continue;
                }

                double value = vals[ch];

                if (double.IsNaN(value))
                {
                    UpdateProbeUI(p, 0, "ERR", false);
                    continue;
                }

                // Pass the ACTUAL raw value directly, no normalization here
                // Change UpdateProbeUI's first argument to raw value
                bool ok = value >= 0.2 && value <= 1.2;

                UpdateProbeUI(p, value, $"{value:0.000}", ok);
            }
        }



        private void UpdateProbeUI(ProbeRow p, double val, string status, bool ok)
        {
            Application.Current.Dispatcher.Invoke(() =>
            {
                p.Value = val;
                p.Status = status;
                p.InRange = ok;
            });
        }

        private string ReadFullResponse(int timeoutMs = 80, CancellationToken? token = null)
        {
            if (_serial == null) return "";
            string resp = "";
            var sw = System.Diagnostics.Stopwatch.StartNew();

            try
            {
                while (sw.ElapsedMilliseconds < timeoutMs)
                {
                    // ✅ CRITICAL: Check cancellation BEFORE reading
                    token?.ThrowIfCancellationRequested();

                    // Only read if data available (non-blocking)
                    if (_serial.BytesToRead > 0)
                    {
                        resp += _serial.ReadExisting();
                        if (resp.Contains("#")) break;
                    }

                    Thread.Sleep(1); // Reduced from 3ms for faster response
                }
            }
            catch (OperationCanceledException)
            {
                return ""; // Graceful exit
            }
            catch { }

            return resp;
        }


        private double[] ParseResponse(string r)
        {
            double[] ch = new double[5];
            for (int i = 1; i <= 4; i++) ch[i] = double.NaN;

            if (string.IsNullOrWhiteSpace(r))
            {
                System.Diagnostics.Debug.WriteLine("❌ Empty response");
                return ch;
            }

            // Remove VALL if present
            int idx = r.IndexOf("VALL", StringComparison.OrdinalIgnoreCase);
            if (idx >= 0) r = r[(idx + 4)..];

            // Clean completely
            r = r.Replace("#", "").Replace("\r", "").Replace("\n", "").Trim();

            // 🌟 GAGENET CSV: C05-01.1814,C06-01.1748,C07-00.0332,C08-00.0315
            var matches = Regex.Matches(r, @"C\d{2}([-+]?\d*\.?\d+)");

            for (int i = 0; i < matches.Count && i < 4; i++)
            {
                string fullMatch = matches[i].Value;        // "C05-01.1814"
                string valueStr = matches[i].Groups[1].Value; // "-01.1814"


                if (double.TryParse(valueStr, NumberStyles.Any, CultureInfo.InvariantCulture, out double v))
                {
                    ch[i + 1] = v;
                    //System.Diagnostics.Debug.WriteLine($"✅ Ch{i + 1}: '{fullMatch}' -> {v:+0.000;-0.000}");
                }
                else
                {
                    // System.Diagnostics.Debug.WriteLine($"❌ Ch{i + 1} Parse failed: '{valueStr}'");
                }
            }

            return ch;
        }

        #endregion

        #region START / STOP MAIN BUTTON

        private Task? _serialReadTask;


        private async void StopBtn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Cancel serial reading task
                _serialCts?.Cancel();

                // Wait up to 500ms for reading task to end gracefully
                if (_serialReadTask != null)
                {
                    await Task.WhenAny(_serialReadTask, Task.Delay(500));
                    _serialReadTask = null;
                }

                _serialCts?.Dispose();
                _serialCts = null;

                // Check if port is open before closing
                if (_serial != null && _serial.IsOpen)
                {
                    _serial.Close();
                    _serial.Dispose();
                    _serial = null;
                }

                System.Diagnostics.Debug.WriteLine("Serial port closed cleanly.");

                // Optionally clear UI or reset probe values here
                foreach (var p in Probes)
                {
                    p.Value = 0;
                    p.Status = "";
                    p.InRange = false;
                }

                if (_timer.IsEnabled)
                    _timer.Stop();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error closing serial port: {ex.Message}");
            }
        }


        private void StartBtn_Click(object sender, RoutedEventArgs e)
        {
            if (_serial == null || !_serial.IsOpen)
            {
                MessageBox.Show("Connect serial first.");
                return;
            }

            if (_serialCts != null)
            {
                MessageBox.Show("Already running.");
                return;
            }

            _serialCts = new CancellationTokenSource();

            // Store the task to await during stop
            _serialReadTask = Task.Run(() => SerialLiveLoopAsync(_serialCts.Token));
        }

        #endregion

        #region IO MONITOR
        private async Task LoadInputDevicesAsync()
        {
            inputDevices.Clear();
            using var conn = new SqlConnection(_connectionString);
            await conn.OpenAsync();
            using var cmd = new SqlCommand("SELECT * FROM IODevices WHERE IsInput = 1", conn);
            using var reader = await cmd.ExecuteReaderAsync();
            while (await reader.ReadAsync())
            {
                inputDevices.Add(new IODevice
                {
                    Description = reader["Description"]?.ToString() ?? "Unknown",
                    Bit = reader["Bit"]?.ToString() ?? "None"
                });
            }
        }

        private void GenerateInputButtons()
        {
            try
            {
                ToggleGrid1.Children.Clear();
                inputButtons.Clear();

                foreach (var device in inputDevices)
                {
                    try
                    {
                        Button inputButton = new Button
                        {
                            Width = 120,
                            Height = 35,
                            Margin = new Thickness(20),
                            Tag = device,
                            Background = Brushes.White,
                            Foreground = Brushes.Black,
                            Opacity = 0.7,
                            FontWeight = FontWeights.Bold,
                            Content = new TextBlock
                            {
                                Text = device.Description,
                                TextWrapping = TextWrapping.Wrap,
                                TextAlignment = TextAlignment.Center,
                                HorizontalAlignment = HorizontalAlignment.Center,
                                VerticalAlignment = VerticalAlignment.Center
                            }
                        };


                        try
                        {
                            inputButton.Style = (Style)FindResource("ModernButtonStyle");
                        }
                        catch (ResourceReferenceKeyNotFoundException)
                        {
                            // Style not found, continue with default
                        }

                        inputButtons[device.Bit] = inputButton;
                        ToggleGrid1.Children.Add(inputButton);
                    }
                    catch (Exception buttonEx)
                    {
                        MessageBox.Show($"Error creating input button for {device.Bit}: {buttonEx.Message}",
                            "Button Creation Error", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error generating input buttons: {ex.Message}", "UI Error",
                    System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
            }
        }

        private void GenerateOutputButtons()
        {
            try
            {
                ToggleGrid.Children.Clear();
                outputButtons.Clear();

                foreach (var device in outputDevices)
                {
                    try
                    {
                        StackPanel container = new StackPanel
                        {
                            Orientation = Orientation.Vertical,
                            Margin = new Thickness(3),
                            HorizontalAlignment = HorizontalAlignment.Center,
                            VerticalAlignment = VerticalAlignment.Center

                        };

                        ToggleButton outputButton = new ToggleButton
                        {
                            Tag = device,
                            IsEnabled = false,
                            HorizontalAlignment = HorizontalAlignment.Center,
                            Margin = new Thickness(0, 0, 0, 5)
                        };

                        try
                        {
                            outputButton.Style = (Style)FindResource("ModernSwitchToggleStyle");
                        }
                        catch (ResourceReferenceKeyNotFoundException)
                        {
                            // Style not found, set default properties
                            outputButton.Width = 80;
                            outputButton.Height = 36;
                        }

                        TextBlock deviceLabel = new TextBlock
                        {
                            Text = $"{device.Description}",
                            Foreground = Brushes.White,
                            FontWeight = FontWeights.Bold,
                            FontSize = 12,
                            TextAlignment = TextAlignment.Center,
                            HorizontalAlignment = HorizontalAlignment.Center,
                            VerticalAlignment = VerticalAlignment.Center,
                            Margin = new Thickness(0, 0, 0, 0)
                        };

                        outputButton.Checked += OutputButton_Checked;
                        outputButton.Unchecked += OutputButton_Unchecked;
                        outputButtons[device.Bit] = outputButton;

                        container.Children.Add(outputButton);
                        container.Children.Add(deviceLabel);
                        ToggleGrid.Children.Add(container);
                    }
                    catch (Exception buttonEx)
                    {
                        MessageBox.Show($"Error creating output button for {device.Bit}: {buttonEx.Message}",
                            "Button Creation Error", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error generating output buttons: {ex.Message}", "UI Error",
                    System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
            }
        }

        private void OutputButton_Checked(object sender, RoutedEventArgs e)
        {
            if (sender is not ToggleButton button || button.Tag is not IODevice device)
                return;

            if (!isPlcConnected)
            {
                MessageBox.Show("PLC not connected.");
                button.IsChecked = false;
                return;
            }

            try
            {
                int result = plc.SetDevice(device.Bit, 1);
                if (result == 0)
                {
                    button.Background = Brushes.Green;
                }
                else
                {
                    button.IsChecked = false;
                    button.Background = Brushes.Red;
                    MessageBox.Show($"Failed to turn {device.Bit} ON. Error code {result}.");
                }
            }
            catch (Exception ex)
            {
                button.IsChecked = false;
                MessageBox.Show($"PLC error for {device.Bit}: {ex.Message}");
            }
        }

        private void OutputButton_Unchecked(object sender, RoutedEventArgs e)
        {
            if (sender is not ToggleButton button || button.Tag is not IODevice device)
                return;

            if (!isPlcConnected)
            {
                MessageBox.Show("PLC not connected.");
                button.IsChecked = true; // Keep ON if not connected
                return;
            }

            try
            {
                int result = plc.SetDevice(device.Bit, 0);
                if (result == 0)
                {
                    button.Background = Brushes.Red;
                }
                else
                {
                    button.IsChecked = true;  // Stay ON on failure
                    button.Background = Brushes.Green;
                    MessageBox.Show($"Failed to turn {device.Bit} OFF. Error: {result}");
                }
            }
            catch (Exception ex)
            {
                button.IsChecked = true;
                MessageBox.Show($"PLC error {device.Bit}: {ex.Message}");
            }
        }



        private async void ReadingToggle_Checked(object sender, RoutedEventArgs e)
        {
            // Same logic as StartBtn_Click
            if (_serial == null || !_serial.IsOpen)
            {
                MessageBox.Show("Connect serial first.");
                ReadingToggleButton.IsChecked = false;
                return;
            }

            if (_serialCts != null)
            {
                MessageBox.Show("Already running.");
                ReadingToggleButton.IsChecked = false;
                return;
            }

            ReadingToggleButton.Content = "Stop Reading";
            _serialCts = new CancellationTokenSource();
            _serialReadTask = Task.Run(() => SerialLiveLoopAsync(_serialCts.Token));
        }

        private void ReadingToggle_Unchecked(object sender, RoutedEventArgs e)
        {
            // Same logic as StopBtn_Click
            try
            {
                _serialCts?.Cancel();

                if (_serialReadTask != null)
                {
                    _ = Task.WhenAny(_serialReadTask, Task.Delay(500));
                    _serialReadTask = null;
                }

                _serialCts?.Dispose();
                _serialCts = null;

                if (_serial != null && _serial.IsOpen)
                {
                    _serial.Close();
                    _serial.Dispose();
                    _serial = null;
                }

                ReadingToggleButton.Content = "Start Reading";

                foreach (var p in Probes)
                {
                    p.Value = 0;
                    p.Status = "";
                    p.InRange = false;
                }

                if (_timer.IsEnabled)
                    _timer.Stop();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error closing serial port: {ex.Message}");
            }
        }



        private void UpdateInputButtonStates()
        {
            if (!isPlcConnected) return;

            foreach (var dev in inputDevices)
            {
                if (!inputButtons.ContainsKey(dev.Bit)) continue;

                int result = plc.GetDevice(dev.Bit, out int value);

                inputButtons[dev.Bit].Background =
                    result == 0
                        ? (value == 1 ? Brushes.Green : Brushes.Red)
                        : Brushes.Gray;
            }
        }

        private void StartIOMonitoring()
        {
            ioMonitorTimer.Tick += (s, e) => UpdateInputButtonStates();
            ioMonitorTimer.Start();
        }

        private void StopIOMonitoring() => ioMonitorTimer.Stop();
        #endregion

        #region CLEANUP
        private void NotifyStatus(string msg) => StatusMessageChanged?.Invoke(msg);

        public void HandleEscKeyAction()
        {
            try
            {
                foreach (var btn in outputButtons.Values)
                    btn.IsEnabled = false;

                _timer?.Stop();
                StopSerial();
                StopIOMonitoring();
                plc?.Close();

                Window window = Window.GetWindow(this);
                var grid = window?.FindName("MainContentGrid") as Grid;
                if (grid != null)
                {
                    grid.Children.Clear();
                    grid.Children.Add(new Dashboard());
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            catch { }
        }
        #endregion

        #region MODELS
        public event PropertyChangedEventHandler? PropertyChanged;
        private void RaisePropertyChanged([CallerMemberName] string? p = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(p));

        public class ProbeRow : INotifyPropertyChanged
        {
            private string _title = "";
            private int _boxId;
            private List<int> _channels = new();
            private string _status = "";
            private double _value;
            private bool _inRange;
            private string _id = string.Empty;


            public string ProbeName { get; set; }   // 🔥 Add this!
            public string ParameterName { get; set; }

            public string ID { get => _id; set => Set(ref _id, value); }

            public string Title { get => _title; set => Set(ref _title, value); }
            public int BoxId { get => _boxId; set => Set(ref _boxId, value); }
            public List<int> Channels { get => _channels; set => Set(ref _channels, value); }
            public string Status { get => _status; set => Set(ref _status, value); }
            public double Value { get => _value; set => Set(ref _value, value); }
            public bool InRange { get => _inRange; set => Set(ref _inRange, value); }

            public event PropertyChangedEventHandler? PropertyChanged;
            private void Set<T>(ref T field, T v, [CallerMemberName] string n = "")
            {
                if (!EqualityComparer<T>.Default.Equals(field, v))
                {
                    field = v;
                    PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(n));
                }
            }
        }

        public class IODevice
        {
            public string Description { get; set; } = "";
            public string Bit { get; set; } = "";
        }
        #endregion
    }
}
