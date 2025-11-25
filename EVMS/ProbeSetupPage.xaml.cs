using ActUtlType64Lib;
using Microsoft.Data.SqlClient;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Configuration;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using static EVMS.IO_Controle_page;

namespace EVMS
{
    public partial class ProbeSetupPage : UserControl, INotifyPropertyChanged
    {
        private readonly string _connectionString;
        private readonly DispatcherTimer _timer;
        private readonly OrbitService _orbitService = new OrbitService();
        public event Action<string>? StatusMessageChanged;

        private IActUtlType64 plc = new ActUtlType64Class();
        private bool isPlcConnected = false;
        private bool plcBitState = false; // Track the toggle state
        private bool motorRunState = false; // Track current run state

        private List<IODevice> inputDevices = new();
        private Dictionary<string, Button> inputButtons = new();
        //private ActUtlType plcDevice = new ActUtlType();
        private readonly DispatcherTimer ioMonitorTimer = new() { Interval = TimeSpan.FromMilliseconds(250) };




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
                    _ = LoadProbeNamesFromPartConfigAsync(_selectedPartNo);
                }
            }
        }

        public ProbeSetupPage()
        {
            InitializeComponent();
            DataContext = this;
            _connectionString = ConfigurationManager.ConnectionStrings["EVMSDb"].ConnectionString;
            _timer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(10) };
            _timer.Tick += Timer_Tick;

            this.Loaded += SettingsPage_Loaded;
            this.PreviewKeyDown += SettingsPage_PreviewKeyDown;
        }

        // 🔹 ESC key to go back to HomePage
        private void SettingsPage_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                HandleEscKeyAction();
                e.Handled = true;
            }
        }

        // 🔹 Page load: connect to Orbit and load part numbers
        // 🔹 Page load: connect to Orbit and load part numbers (non-blocking)
        private async void SettingsPage_Loaded(object? sender, RoutedEventArgs e)
        {
            try
            {
                await LoadInputDevicesAsync();
                GenerateInputButtons();
                StartIOMonitoring();

                this.Focusable = true;
                this.IsTabStop = true;
                Keyboard.Focus(this);
                FocusManager.SetFocusedElement(Window.GetWindow(this)!, this);

                // Load part numbers
                await LoadPartNumbersAsync();   // ← auto selected first one

                // Auto-load probe names + details
                if (!string.IsNullOrEmpty(SelectedPartNo))
                {
                    await LoadProbeNamesFromPartConfigAsync(SelectedPartNo);
                    await LoadProbeDetailsFromInstallationDataAsync(SelectedPartNo);
                }

                // Start Orbit connection in background
                _ = Task.Run(async () =>
                {
                    NotifyStatus("Connecting to Orbit...");

                    bool connected = await _orbitService.ConnectAsync();

                    Application.Current.Dispatcher.Invoke(async () =>
                    {
                        if (connected)
                        {
                            NotifyStatus($"Connected to Probes");

                            // Once Orbit is connected → load again + start progress bar
                            if (!string.IsNullOrEmpty(SelectedPartNo))
                            {
                                await LoadProbeNamesFromPartConfigAsync(SelectedPartNo);
                                await LoadProbeDetailsFromInstallationDataAsync(SelectedPartNo);
                            }

                            // Auto-start live reading
                            _timer.Start();
                        }
                        else
                        {
                            MessageBox.Show("Failed to connect Orbit.");
                        }
                    });
                });

                ConnectPlcOnPageLoad();
            }
            catch (Exception ex)
            {
                NotifyStatus($"Error: {ex.Message}");
            }
        }


        // 🔹 PLC Auto-Connect Method
        // 🔹 PLC Auto-Connect (Async)
        // 🔹 Asynchronous PLC connection (background, non-blocking)
        private void ConnectPlcOnPageLoad()
        {
            NotifyStatus("🔄 Connecting to PLC...");

            //Task.Run(async () =>
            //{
            //    try
            //    {
            //        plc.ActLogicalStationNumber = 1; // 🧩 Your station number
            //        int openResult = -1;

            //        // Attempt up to 3 retries (optional)
            //        for (int attempt = 1; attempt <= 3; attempt++)
            //        {
            //            try
            //            {
            //                openResult = plc.Open();
            //            }
            //            catch
            //            {
            //                openResult = -1; // Mark as failed
            //            }

            //            if (openResult == 0)
            //                break; // success, exit retry loop

            //            await Task.Delay(1000); // small pause before retry
            //        }

            //        // Back to UI thread to update status
            //        Application.Current.Dispatcher.Invoke(() =>
            //        {
            //            if (openResult == 0)
            //            {
            //                isPlcConnected = true;
            //                NotifyStatus("✅ PLC connected successfully.");
            //            }
            //            else
            //            {
            //                isPlcConnected = false;
            //                NotifyStatus("⚠️ PLC connection failed after retries.");
            //                MessageBox.Show("⚠️ Failed to connect to PLC. Please check the connection.",
            //                    "PLC Connection Error", MessageBoxButton.OK, MessageBoxImage.Warning);
            //            }
            //              });
            //        }
            //    catch (Exception ex)
            //    {
            //        Application.Current.Dispatcher.Invoke(() =>
            //        {
            //            isPlcConnected = false;
            //            NotifyStatus($"❌ Error connecting to PLC: {ex.Message}");
            //            MessageBox.Show($"Error connecting to PLC: {ex.Message}",
            //                "PLC Error", MessageBoxButton.OK, MessageBoxImage.Error);
            //        });
            //    }
            //});
        }




        // 🔹 Motor Up/Down Toggle Button
        private void MotorUpDownButton_Click(object sender, RoutedEventArgs e)
        {
            if (!isPlcConnected)
            {
                MessageBox.Show("⚠️ PLC not connected. Please check connection.",
                    "PLC Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            string plcBit = "M10"; // PLC bit for Motor Up/Down

            try
            {
                // Toggle state
                plcBitState = !plcBitState;
                int valueToWrite = plcBitState ? 1 : 0;

                int result = plc.SetDevice(plcBit, valueToWrite);

                if (result == 0)
                {
                    MotorUpDownButton.Content = plcBitState ? "Cylinder: DOWN" : "Cylinder: UP";
                    NotifyStatus($"PLC bit {plcBit} set to {valueToWrite} (Motor {(plcBitState ? "Started" : "Stopped")})");
                }
                else
                {
                    MessageBox.Show($"Failed to write to {plcBit}. Error Code: {result}",
                        "PLC Write Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error toggling Motor bit: {ex.Message}",
                    "PLC Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void MotorRunButton_Click(object sender, RoutedEventArgs e)
        {
            if (!isPlcConnected)
            {
                MessageBox.Show("⚠️ PLC not connected. Please check connection.",
                    "PLC Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            string plcBit = "M14"; // 🔹 PLC bit for Motor Run (different from M10)

            try
            {
                // Toggle state
                motorRunState = !motorRunState;
                int valueToWrite = motorRunState ? 1 : 0;

                int result = plc.SetDevice(plcBit, valueToWrite);

                if (result == 0)
                {
                    MotorRunButton.Content = motorRunState ? "MOTOR: ON" : "MOTOR: OFF";
                    NotifyStatus($"PLC bit {plcBit} set to {valueToWrite} (Motor {(motorRunState ? "Running" : "Stopped")})");
                }
                else
                {
                    MessageBox.Show($"Failed to write to {plcBit}. Error Code: {result}",
                        "PLC Write Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error toggling Motor Run bit: {ex.Message}",
                    "PLC Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }


        // 🔹 Load Part Numbers
        private async Task LoadPartNumbersAsync()
        {
            try
            {
                using var con = new SqlConnection(_connectionString);
                await con.OpenAsync();
                using var cmd = new SqlCommand("SELECT DISTINCT Para_No FROM PART_ENTRY ORDER BY Para_No", con);
                using var reader = await cmd.ExecuteReaderAsync();
                PartNumbers.Clear();
                while (await reader.ReadAsync())
                {
                    PartNumbers.Add(reader.GetString(0));
                }

                if (PartNumbers.Count > 0)
                {
                    SelectedPartNo = PartNumbers[0];
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading Part Numbers: {ex.Message}", "Database Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // 🔹 Load Probe Names from Config
        private async Task LoadProbeNamesFromPartConfigAsync(string partNo)
        {
            if (string.IsNullOrEmpty(partNo)) return;
            try
            {
                using var con = new SqlConnection(_connectionString);
                await con.OpenAsync();
                const string query = @"SELECT Parameter,D_Name FROM PartConfig WHERE Para_No = @Para_No AND ProbeStatus = 'Probe'";
                using var cmd = new SqlCommand(query, con);
                cmd.Parameters.AddWithValue("@Para_No", partNo);
                using var reader = await cmd.ExecuteReaderAsync();
                Probes.Clear();
                while (await reader.ReadAsync())
                {
                    Probes.Add(new ProbeRow
                    {
                        Title = reader.GetString(1),
                        ID = string.Empty,
                        Stroke = string.Empty,
                        Value = 0,
                        InRange = false
                    });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading probe names: {ex.Message}", "Database Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // 🔹 Load Probe Installation Details
        private async Task LoadProbeDetailsFromInstallationDataAsync(string partNo)
        {
            if (string.IsNullOrEmpty(partNo)) return;
            try
            {
                using var con = new SqlConnection(_connectionString);
                await con.OpenAsync();
                const string query = @"SELECT ProbeId, Stroke 
                                       FROM ProbeInstallationData 
                                       WHERE PartNo = @PartNo AND Status = @Status";
                using var cmd = new SqlCommand(query, con);
                cmd.Parameters.AddWithValue("@PartNo", partNo);
                cmd.Parameters.AddWithValue("@Status", "Installed");
                using var reader = await cmd.ExecuteReaderAsync();

                int index = 0;
                while (await reader.ReadAsync() && index < Probes.Count)
                {
                    string probeId = reader.GetString(0);
                    string stroke = reader.GetString(1);

                    var probe = Probes[index];
                    if (_orbitService.IsModuleConnected(probeId))
                    {
                        probe.ID = probeId;
                        probe.Stroke = stroke;
                    }
                    else
                    {
                        probe.ID = string.Empty;
                        probe.Stroke = string.Empty;
                    }
                    index++;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading probe details: {ex.Message}", "Database Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // 🔹 Live Reading Timer
        private void Timer_Tick(object? sender, EventArgs? e)
        {
            if (_orbitService.ModulesById == null || _orbitService.ModulesById.Count == 0)
                return;

            double maxScale = 2.0;

            foreach (var probe in Probes)
            {
                if (probe == null || string.IsNullOrEmpty(probe.ID))
                    continue;

                if (!_orbitService.ModulesById.TryGetValue(probe.ID, out dynamic module))
                    continue;

                try
                {
                    double reading = (double)module.ReadingInUnits;
                    double normalizedValue = (reading / maxScale) * 100;
                    probe.Value = Math.Min(Math.Max(normalizedValue, 0), 100);

                    if (reading > 1.200)
                    {
                        probe.Status = "OVER";
                        probe.InRange = false;
                    }
                    else if (reading < 0.2)
                    {
                        probe.Status = "UNDER";
                        probe.InRange = false;
                    }
                    else
                    {
                        probe.Status = $"{reading:0.000} mm";
                        probe.InRange = true;
                    }
                }
                catch
                {
                    probe.Value = 0;
                    probe.Status = "ERR";
                    probe.InRange = false;
                }
            }
        }

        // 🔹 Start Button
        private async void StartBtn_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(SelectedPartNo))
            {
                MessageBox.Show("Please select a Part Number first.", "Input Required", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (!_orbitService.IsConnected)
            {
                MessageBox.Show("Orbit Controller is not connected. Please reconnect and try again.",
                    "Connection Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            await LoadProbeNamesFromPartConfigAsync(SelectedPartNo);
            await LoadProbeDetailsFromInstallationDataAsync(SelectedPartNo);

            var missingIds = new List<string>();
            foreach (var probe in Probes)
            {
                if (probe == null || string.IsNullOrEmpty(probe.ID))
                    missingIds.Add(probe?.Title ?? "<Unknown>");
            }

            if (missingIds.Count > 0)
            {
                string missing = string.Join(", ", missingIds);
                MessageBox.Show($"The following probes are missing or not connected: {missing}",
                    "Probe Missing", MessageBoxButton.OK, MessageBoxImage.Warning);
            }

            _timer.Start();
            //MessageBox.Show("Live probe reading started.",
            //    "Reading Active", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        // 🔹 Stop Button
        private void StopBtn_Click(object sender, RoutedEventArgs e)
        {
            _timer.Stop();
            foreach (var probe in Probes)
            {
                probe.Value = 0;
                probe.ID = string.Empty;
                probe.Stroke = string.Empty;
                probe.Status = string.Empty;
                probe.InRange = false;
            }
            //_orbitService.Disconnect();
            //MessageBox.Show("Live reading stopped.", "Stopped", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        // 🔹 PropertyChanged boilerplate
        public event PropertyChangedEventHandler? PropertyChanged;
        private void RaisePropertyChanged([CallerMemberName] string? propertyName = null) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));

        // 🔹 Inner Class for ProbeRow
        public class ProbeRow : INotifyPropertyChanged
        {
            private string _title = string.Empty;
            private string _id = string.Empty;
            private string _stroke = string.Empty;
            private string _status = string.Empty;
            private double _value;
            private bool _inRange;

            public string Title { get => _title; set => SetField(ref _title, value); }
            public string ID { get => _id; set => SetField(ref _id, value); }
            public string Stroke { get => _stroke; set => SetField(ref _stroke, value); }
            public string Status { get => _status; set => SetField(ref _status, value); }
            public double Value { get => _value; set => SetField(ref _value, value); }
            public bool InRange { get => _inRange; set => SetField(ref _inRange, value); }

            public event PropertyChangedEventHandler? PropertyChanged;
            protected bool SetField<T>(ref T field, T value, [CallerMemberName] string? propertyName = null)
            {
                if (Equals(field, value)) return false;
                field = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
                return true;
            }
        }
        private void NotifyStatus(string message)
        {
            StatusMessageChanged?.Invoke(message);
        }

        // 🔹 ESC Key Handler (go to Home)
        private void HandleEscKeyAction()
        {
            try
            {
                _timer?.Stop();
                _orbitService?.Disconnect();
                StopIOMonitoring();
                plc?.Close();

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
                        resultPage.Focus();
                    }
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error while cleaning up resources: {ex.Message}",
                    "Cleanup Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

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
            InputButtonPanel.Children.Clear();
            inputButtons.Clear();
            foreach (var device in inputDevices)
            {
                Button inputButton = new Button
                        {
                            Width = 140,
                            Height = 40,
                            Margin = new Thickness(10),
                            Tag = device,
                            Background = Brushes.White,
                            Foreground = Brushes.Black,
                            Opacity=0.7,
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
                InputButtonPanel.Children.Add(inputButton);
            }
        }

        private void UpdateInputButtonStates()
        {
            try
            {
                if (!isPlcConnected) return;

                foreach (var device in inputDevices)
                {
                    try
                    {
                        if (inputButtons.ContainsKey(device.Bit))
                        {
                            int value;
                            int result = plc.GetDevice(device.Bit, out value);
                            if (result == 0)
                            {
                                inputButtons[device.Bit].Background = (value == 1) ? Brushes.Green : Brushes.Red;
                            }
                            else
                            {
                                inputButtons[device.Bit].Background = Brushes.Gray;
                            }
                        }

                    }
                    catch (Exception ex)
                    {
                        // Optionally log ex.Message
                        if (inputButtons.TryGetValue(device.Bit.Trim(), out var inputButton))
                        {
                            inputButton.Background = Brushes.Gray;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Stop timer if critical error, alert user
                ioMonitorTimer.Stop();
                MessageBox.Show($"Input monitor error: {ex.Message}", "Monitor Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }




        private void StartIOMonitoring()
        {
            ioMonitorTimer.Tick += (s, e) => UpdateInputButtonStates();
            ioMonitorTimer.Start();
        }
        private void StopIOMonitoring()
        {
            ioMonitorTimer.Stop();
        }

        public class IODevice
        {
            public string Description { get; set; } = "";
            public string Bit { get; set; } = "";
        }


    }
}
