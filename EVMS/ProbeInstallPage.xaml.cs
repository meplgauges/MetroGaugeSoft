using Microsoft.Data.SqlClient;
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Configuration;
using System.Globalization;
using System.IO.Ports;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
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
        private readonly string connectionString;
        private readonly int[] _pos = { 4, 16, 28, 40 };

        public ObservableCollection<ParameterItem> Parameters { get; set; } = new ObservableCollection<ParameterItem>();

        public ProbeInstallPage()
        {
            InitializeComponent();
            connectionString = ConfigurationManager.ConnectionStrings["EVMSDb"].ConnectionString!;
            DataContext = this;
            
            // AUTO-LOAD ON PAGE INIT
            Loaded += async (s, e) => await InitializePageAsync();
        }

        // ------------------ PAGE INITIALIZATION ------------------  
        private async Task InitializePageAsync()
        {
            await LoadParametersAsync();
            RefreshPorts();
        }

        // ------------------ LOAD PARAMETERS ------------------  
        private async Task LoadParametersAsync()
        {
            Parameters.Clear();
            TxtStatus.Text = "Loading parameters...";

            try
            {
                using var con = new SqlConnection(connectionString);
                await con.OpenAsync();

                string partQuery = "SELECT TOP 1 Para_No FROM Part_Entry WHERE ActivePart = 1";
                string? partNo = (string?)await new SqlCommand(partQuery, con).ExecuteScalarAsync();

                if (partNo == null)
                {
                    TxtStatus.Text = "No active part found.";
                    return;
                }

                string paramQuery = "SELECT Parameter FROM PartConfig WHERE Para_No=@P AND ProbeStatus='Probe'";
                using var cmd = new SqlCommand(paramQuery, con);
                cmd.Parameters.AddWithValue("@P", partNo);

                using var r = await cmd.ExecuteReaderAsync();
                while (await r.ReadAsync())
                    Parameters.Add(new ParameterItem(r.GetString(0)));
                r.Close();

                string mapQuery = "SELECT ParameterName, BoxId, ChannelId FROM ProbeInstallationData WHERE PartNo=@P";
                using var mcmd = new SqlCommand(mapQuery, con);
                mcmd.Parameters.AddWithValue("@P", partNo);

                using var mread = await mcmd.ExecuteReaderAsync();
                while (await mread.ReadAsync())
                {
                    string pname = mread.GetString(0);
                    int box = mread.GetInt32(1);
                    int ch = mread.GetInt32(2);

                    var p = Parameters.FirstOrDefault(x => x.Name == pname);
                    if (p != null)
                    {
                        p.BoxId = box;
                        if (!p.Channels.Contains(ch))
                            p.Channels.Add(ch);
                    }
                }

                TxtStatus.Text = $"Loaded {Parameters.Count} parameters. Assign Box + Channels, then SAVE.";
            }
            catch (Exception ex)
            {
                TxtStatus.Text = "Load error: " + ex.Message;
                MessageBox.Show("Load error: " + ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // ------------------ SAVE MAPPINGS ------------------  
        private async void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            TxtStatus.Text = "Saving...";
            try
            {
                using var con = new SqlConnection(connectionString);
                await con.OpenAsync();

                string partQuery = "SELECT TOP 1 Para_No FROM Part_Entry WHERE ActivePart = 1";
                string? partNo = (string?)await new SqlCommand(partQuery, con).ExecuteScalarAsync();

                if (partNo == null)
                {
                    MessageBox.Show("No active part.");
                    return;
                }

                var del = new SqlCommand("DELETE FROM ProbeInstallationData WHERE PartNo=@P", con);
                del.Parameters.AddWithValue("@P", partNo);
                await del.ExecuteNonQueryAsync();

                int savedCount = 0;
                foreach (var p in Parameters)
                {
                    foreach (int ch in p.Channels)
                    {
                        var ins = new SqlCommand(
                            "INSERT INTO ProbeInstallationData (PartNo,ParameterName,BoxId,ChannelId) VALUES (@p,@n,@b,@c)", con);
                        ins.Parameters.AddWithValue("@p", partNo);
                        ins.Parameters.AddWithValue("@n", p.Name);
                        ins.Parameters.AddWithValue("@b", p.BoxId);
                        ins.Parameters.AddWithValue("@c", ch);
                        await ins.ExecuteNonQueryAsync();
                        savedCount++;
                    }
                }

                TxtStatus.Text = $"Saved {savedCount} probe assignments.";
                MessageBox.Show($"Probe assignments saved successfully ({savedCount} mappings).", "Success");
            }
            catch (Exception ex)
            {
                TxtStatus.Text = "Save error: " + ex.Message;
                MessageBox.Show("Save error: " + ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // ------------------ SERIAL PORT ------------------  
        private void RefreshPorts()
        {
            try
            {
                ComboPorts.ItemsSource = SerialPort.GetPortNames().OrderBy(p => p).ToArray();
                if (ComboPorts.Items.Count > 0 && ComboPorts.SelectedIndex < 0)
                    ComboPorts.SelectedIndex = 0;
                TxtStatus.Text += " | COM ports refreshed.";
            }
            catch (Exception ex)
            {
                TxtStatus.Text = "Refresh ports error: " + ex.Message;
            }
        }

        private void BtnRefreshPorts_Click(object sender, RoutedEventArgs e) => RefreshPorts();

        private void BtnConnect_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (ComboPorts.SelectedItem == null)
                {
                    MessageBox.Show("Select a COM port.");
                    return;
                }

                StopLive(); // Stop any existing live reading
                _serial?.Dispose();
                
                _serial = new SerialPort(ComboPorts.SelectedItem.ToString(), 115200, Parity.None, 8, StopBits.One)
                {
                    ReadTimeout = 500,
                    WriteTimeout = 500
                };
                _serial.Open();
                TxtStatus.Text = $"Connected to {ComboPorts.SelectedItem}.";
            }
            catch (Exception ex)
            {
                TxtStatus.Text = "Connect error: " + ex.Message;
                MessageBox.Show("Failed to open port: " + ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BtnDisconnect_Click(object sender, RoutedEventArgs e)
        {
            StopLive();
            try
            {
                _serial?.Close();
                _serial?.Dispose();
                _serial = null;
                TxtStatus.Text = "Disconnected.";
            }
            catch (Exception ex)
            {
                TxtStatus.Text = "Disconnect error: " + ex.Message;
            }
        }

        // ------------------ LIVE READING ------------------  
        private void BtnStart_Click(object sender, RoutedEventArgs e)
        {
            if (_serial == null || !_serial.IsOpen)
            {
                MessageBox.Show("Connect serial port first.");
                return;
            }

            if (_cts != null)
            {
                MessageBox.Show("Already running.");
                return;
            }

            _cts = new CancellationTokenSource();
            _ = Task.Run(() => LiveLoopAsync(_cts.Token));
            TxtStatus.Text = "🔴 Live reading started (click STOP to pause).";
        }

        private void BtnStop_Click(object sender, RoutedEventArgs e) => StopLive();

        private void StopLive()
        {
            try
            {
                _cts?.Cancel();
                _cts?.Dispose();
                _cts = null;
                TxtStatus.Text = TxtStatus.Text.Replace("🔴 Live reading started", "Live stopped");
            }
            catch (Exception ex)
            {
                TxtStatus.Text = "Stop error: " + ex.Message;
            }
        }

        private async Task LiveLoopAsync(CancellationToken token)
        {
            try
            {
                while (!token.IsCancellationRequested)
                {
                    await ReadOnceAsync();
                    await Task.Delay(8, token);
                }
            }
            catch (OperationCanceledException) { }
            catch (Exception ex)
            {
                Application.Current.Dispatcher.Invoke(() => 
                    TxtStatus.Text = "Live error: " + ex.Message);
            }
        }

        // ------------------ LIVE READING WITH DEBUG ------------------  
        private async Task ReadOnceAsync()
        {
            if (_serial == null || !_serial.IsOpen) return;

            var groups = Parameters.Where(p => p.BoxId > 0).GroupBy(p => p.BoxId);

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
                    foreach (var p in g) UpdateParamValue(p, "ERR");
                    continue;
                }

                // INCREASED TIMEOUT FOR BETTER DATA CAPTURE
                string resp = ReadFullResponse(150);  // 150ms timeout
                if (string.IsNullOrEmpty(resp))
                {
                    foreach (var p in g) UpdateParamValue(p, "ERR");
                    continue;
                }

                // 🚨 DEBUG: SHOW RAW SERIAL DATA
             

                double[] vals = ParseResponse(resp);

                // 🚨 DEBUG: SHOW PARSED VALUES
              

                foreach (var p in g)
                {
                    var assigned = p.Channels.Where(ch => ch >= 1 && ch <= 4).ToList();

                    if (assigned.Count == 0)
                    {
                        UpdateParamValue(p, "-");
                        continue;
                    }

                    var numericValues = assigned.Select(ch => vals[ch]).ToList();
                    bool anyInvalid = numericValues.Any(v => double.IsNaN(v));

                    if (anyInvalid)
                    {
                        UpdateParamValue(p, "ERR");
                        continue;
                    }

                    double sum = numericValues.Sum();
                    // ALWAYS SHOW SIGN +0.123 or -0.456
                    string formatted = sum.ToString("+0.000;-0.000;+0.000", CultureInfo.InvariantCulture);
                    UpdateParamValue(p, formatted);
                }
            }
        }

        // ------------------ ENHANCED SERIAL RESPONSE READER ------------------  
        private string ReadFullResponse(int timeoutMs = 150)
        {
            if (_serial == null) return string.Empty;

            var buffer = new System.Text.StringBuilder();
            var sw = System.Diagnostics.Stopwatch.StartNew();

            try
            {
                while (sw.ElapsedMilliseconds < timeoutMs)
                {
                    if (_serial.BytesToRead > 0)
                    {
                        string chunk = _serial.ReadExisting();
                        buffer.Append(chunk);
                    }

                    string resp = buffer.ToString();
                    if (resp.Contains("#"))
                    {
                        // DEBUG: LOG COMPLETE RAW RESPONSE TO VS OUTPUT
                        return resp;
                    }

                    Thread.Sleep(3);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Read error: {ex.Message}");
            }

            string finalResp = buffer.ToString();
            return finalResp;
        }

        // ------------------ ENHANCED PARSER WITH DEBUG ------------------  
        private double[] ParseResponse(string r)
        {
            double[] ch = new double[5];
            for (int i = 1; i <= 4; i++) ch[i] = double.NaN;

            if (string.IsNullOrWhiteSpace(r))
            {
                return ch;
            }

            // Remove VALL if present
            int idx = r.IndexOf("VALL", StringComparison.OrdinalIgnoreCase);
            if (idx >= 0 && idx + 4 < r.Length)
                r = r.Substring(idx + 4);

            // Clean
            r = r.Replace("#", "").Replace("\r", "").Replace("\n", "").Trim();
            

            // 🌟 GAGENET CSV FORMAT: C05-01.1814,C06-01.1748,...
            var matches = Regex.Matches(r, @"C\d{2}([-+]?\d*\.?\d+)");

         

            for (int i = 0; i < matches.Count && i < 4; i++)
            {
                string fullMatch = matches[i].Value;        // "C05-01.1814"
                string valueStr = matches[i].Groups[1].Value; // "-01.1814"

              

                if (double.TryParse(valueStr, NumberStyles.Float, CultureInfo.InvariantCulture, out double v))
                {
                    ch[i + 1] = v;
                    
                }
                else
                {
                    //System.Diagnostics.Debug.WriteLine($"❌ Ch{i + 1} Parse failed: '{valueStr}'");
                }
            }

            return ch;
        }



        private string ExtractFirstNumber(string s)
        {
            if (string.IsNullOrWhiteSpace(s))
                return string.Empty;

            // Accepts: -1.234, +0.55, .234, 123., 0.123
            var m = Regex.Match(s, @"[-+]?(?:\d*\.\d+|\d+\.?\d*)");

            return m.Success ? m.Value : string.Empty;
        }



        private void UpdateParamValue(ParameterItem p, string text)
        {
            Application.Current.Dispatcher.Invoke(() => p.LiveValue = text);
        }
    }

    // ParameterItem class remains the same
    public class ParameterItem : INotifyPropertyChanged
    {
        private string _name = "";
        private int _boxId;
        private string _liveValue = "-";

        public string Name { get => _name; set { _name = value; Notify(); } }
        public int BoxId { get => _boxId; set { _boxId = value; Notify(); } }
        public ObservableCollection<int> Channels { get; set; } = new();

        public string ChannelsAsString
        {
            get => string.Join(",", Channels);
            set
            {
                Channels.Clear();
                if (!string.IsNullOrWhiteSpace(value))
                    foreach (var p in value.Split(',', ';', ' '))
                        if (int.TryParse(p, out int ch) && ch >= 1 && ch <= 4) Channels.Add(ch);
                Notify();
            }
        }

        public string LiveValue { get => _liveValue; set { _liveValue = value; Notify(); } }

        public ParameterItem(string name) => Name = name;
        public ParameterItem() { }

        public event PropertyChangedEventHandler? PropertyChanged;
        private void Notify([CallerMemberName] string prop = "") => 
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(prop));
    }
}
