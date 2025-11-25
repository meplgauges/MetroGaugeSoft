using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.IO.Ports;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;

namespace EVMS
{
    public partial class ProbeInstallPage : UserControl
    {
        private SerialPort? _serial;
        private CancellationTokenSource? _cts;
        private Task? _liveTask;

        // channel text positions in VALL response (as in your console log)
        private readonly int[] _positions = { 4, 16, 28, 40 }; // start positions (1-based)
        private const int ChannelCount = 4;

        public ObservableCollection<ParameterItem> Parameters { get; set; } = new ObservableCollection<ParameterItem>();

        public ProbeInstallPage()
        {
            InitializeComponent();
            DataContext = this;

            // sample parameters (edit/add in grid)
            Parameters.Add(new ParameterItem("ReducedDia", 1, 1, 3));   // channels 1 & 3 from box 1
            Parameters.Add(new ParameterItem("Length", 2, 2, 4));       // channels 2 & 4 from box 2
            Parameters.Add(new ParameterItem("Width", 1, 1));           // single channel
            Parameters.Add(new ParameterItem("Ovality", 3, 1, 2, 3));   // channels 1,2,3 from box 3
        }

        // ---------- UI lifecycle ----------
        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            RefreshPorts();
            GridParameters.ItemsSource = Parameters;
            TxtStatus.Text = "Ready.";
        }

        private void BtnRefreshPorts_Click(object sender, RoutedEventArgs e) => RefreshPorts();

        private void RefreshPorts()
        {
            try
            {
                var ports = SerialPort.GetPortNames().OrderBy(p => p).ToArray();
                ComboPorts.ItemsSource = ports;
                if (ports.Length > 0) ComboPorts.SelectedIndex = 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error enumerating ports: " + ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // ---------- Connect / Disconnect ----------
        private void BtnConnect_Click(object sender, RoutedEventArgs e)
        {
            if (ComboPorts.SelectedItem == null)
            {
                MessageBox.Show("Select a COM port.", "Info", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            try
            {
                _serial?.Close();
                _serial = new SerialPort(ComboPorts.SelectedItem.ToString()!, 115200, Parity.None, 8, StopBits.One)
                {
                    ReadTimeout = 500,
                    WriteTimeout = 500
                };
                _serial.Open();
                TxtStatus.Text = $"Connected to {ComboPorts.SelectedItem}";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to open port: " + ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BtnDisconnect_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                StopLive();
                _serial?.Close();
                _serial = null;
                TxtStatus.Text = "Disconnected.";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Disconnect error: " + ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // ---------- Start / Stop live ----------
        private void BtnStart_Click(object sender, RoutedEventArgs e)
        {
            if (_serial == null || !_serial.IsOpen)
            {
                MessageBox.Show("Connect serial port first.", "Info", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            if (_cts != null)
            {
                MessageBox.Show("Already running.", "Info", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            _cts = new CancellationTokenSource();
            _liveTask = Task.Run(() => LiveLoopAsync(_cts.Token));
            TxtStatus.Text = "Live reading started.";
        }

        private void BtnStop_Click(object sender, RoutedEventArgs e) => StopLive();

        private void StopLive()
        {
            try
            {
                _cts?.Cancel();
                _cts = null;
                _liveTask = null;
                TxtStatus.Text = "Live reading stopped.";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error stopping: " + ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // ---------- Live loop ----------
        private async Task LiveLoopAsync(CancellationToken token)
        {
            try
            {
                while (!token.IsCancellationRequested)
                {
                    await ReadAllParametersOnceAsync();

                    // small delay (adjust as needed) - 10 Hz would be 100ms; your code used 1 ms earlier, choose appropriate.
                    await Task.Delay(1, token);
                }
            }
            catch (OperationCanceledException) { }
            catch (Exception ex)
            {
                Application.Current.Dispatcher.Invoke(() =>
                {
                    TxtStatus.Text = "Live error: " + ex.Message;
                });
            }
        }

        // ---------- Read grouped by box ----------
        private async Task ReadAllParametersOnceAsync()
        {
            if (_serial == null || !_serial.IsOpen) return;

            // group parameters by BoxId (parameters with same BoxId are read with single VALL)
            var groups = Parameters
                .Where(p => p.BoxId > 0)
                .GroupBy(p => p.BoxId)
                .ToList();

            foreach (var g in groups)
            {
                int box = g.Key;
                string cmd = $"*{box:D3}VALL#\r";

                try
                {
                    _serial.DiscardInBuffer();
                    _serial.DiscardOutBuffer();
                    _serial.Write(cmd);
                }
                catch
                {
                    // on write error -> mark group's values as ERR
                    foreach (var p in g)
                        UpdateParamValue(p, "ERR");
                    continue;
                }

                // wait for response (tune as per device)
                await Task.Delay(90);

                string resp;
                try
                {
                    resp = _serial.ReadExisting();
                }
                catch
                {
                    resp = string.Empty;
                }

                // parse the 4 channel values
                double[] channels = ParseVallResponse(resp);

                // map to parameters in this box
                foreach (var param in g)
                {
                    // Collect values for all channels of this parameter
                    var valuesText = param.Channels
                        .Select(ch =>
                        {
                            if (ch >= 1 && ch <= ChannelCount)
                            {
                                double v = channels[ch];
                                return double.IsNaN(v) ? "ERR" : v.ToString("0.0000", CultureInfo.InvariantCulture);
                            }
                            else
                            {
                                return "NoCh";
                            }
                        })
                        .ToArray();

                    string final = string.Join(", ", valuesText);
                    UpdateParamValue(param, final);
                }
            }
        }

        // ---------- parse VALL response into channels[1..4] ----------
        private double[] ParseVallResponse(string response)
        {
            double[] readings = new double[ChannelCount + 1]; // 1-based
            for (int i = 1; i <= ChannelCount; i++) readings[i] = double.NaN;

            if (string.IsNullOrWhiteSpace(response) || response.Length < _positions.Last())
                return readings;

            for (int ch = 1; ch <= ChannelCount; ch++)
            {
                int start = _positions[ch - 1] - 1; // convert to 0-based
                int len = 8;
                if (start + len <= response.Length)
                {
                    string seg = response.Substring(start, len).Trim();

                    if (double.TryParse(seg, System.Globalization.NumberStyles.Any, CultureInfo.InvariantCulture, out double val))
                    {
                        readings[ch] = val;
                    }
                    else
                    {
                        // try to extract numeric part
                        string num = ExtractFirstNumber(seg);
                        if (double.TryParse(num, System.Globalization.NumberStyles.Any, CultureInfo.InvariantCulture, out val))
                            readings[ch] = val;
                        else
                            readings[ch] = double.NaN;
                    }
                }
            }

            return readings;
        }

        // extract first numeric from a string (e.g., "+012.345V" -> "012.345")
        private string ExtractFirstNumber(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return string.Empty;
            int start = -1, end = -1;
            for (int i = 0; i < s.Length; i++)
            {
                if ((char.IsDigit(s[i]) || s[i] == '+' || s[i] == '-' || s[i] == '.') && start == -1) start = i;
                if (start != -1 && !(char.IsDigit(s[i]) || s[i] == '+' || s[i] == '-' || s[i] == '.'))
                {
                    end = i - 1;
                    break;
                }
            }
            if (start == -1) return string.Empty;
            if (end == -1) end = s.Length - 1;
            return s.Substring(start, end - start + 1);
        }

        // update param's LiveValue on UI thread
        private void UpdateParamValue(ParameterItem p, string text)
        {
            Application.Current.Dispatcher.Invoke(() => p.LiveValue = text);
        }
    }

    // ---------------------------------------------------------
    // Parameter model with INotifyPropertyChanged
    // ---------------------------------------------------------
    public class ParameterItem : INotifyPropertyChanged
    {
        private string _name = "";
        private int _boxId;
        private string _liveValue = "-";

        public string Name
        {
            get => _name;
            set { _name = value; Notify(); }
        }

        // one box per parameter
        public int BoxId
        {
            get => _boxId;
            set { _boxId = value; Notify(); }
        }

        // list of channels (ex: 1,3)
        public ObservableCollection<int> Channels { get; set; } = new ObservableCollection<int>();

        // convenience string for editing channels in DataGrid
        public string ChannelsAsString
        {
            get => string.Join(",", Channels);
            set
            {
                Channels.Clear();
                if (!string.IsNullOrWhiteSpace(value))
                {
                    var parts = value.Split(new[] { ',', ';', ' ' }, StringSplitOptions.RemoveEmptyEntries);
                    foreach (var p in parts)
                    {
                        if (int.TryParse(p.Trim(), out int ch))
                        {
                            Channels.Add(ch);
                        }
                    }
                }
                Notify(nameof(ChannelsAsString));
                Notify(nameof(Channels));
            }
        }

        public string LiveValue
        {
            get => _liveValue;
            set { _liveValue = value; Notify(nameof(LiveValue)); }
        }

        public ParameterItem() { }

        public ParameterItem(string name, int boxId, params int[] channels)
        {
            _name = name;
            _boxId = boxId;
            foreach (var c in channels) Channels.Add(c);
            _liveValue = "-";
        }

        public event PropertyChangedEventHandler? PropertyChanged;
        private void Notify([CallerMemberName] string? propertyName = null) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}
