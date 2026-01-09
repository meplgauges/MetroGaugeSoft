using ActUtlType64Lib;
using Microsoft.Data.SqlClient;
using System.Collections.Concurrent;
using System.Configuration;
using System.Globalization;
using System.IO.Ports;
using System.Text.RegularExpressions;
using System.Threading;

internal class PlcProbeService : IDisposable
{
    private IActUtlType64 plc;
    private SerialPort? _serial;
    private bool isPlcConnected = false;
    private bool isSerialConnected = false;
    private CancellationTokenSource? _cts;
    private Task? _readTask;
    private readonly string _connectionString;

    // 🔥 ORIGINAL READ-ONLY COLLECTIONS (kept for compatibility)
    public IReadOnlyDictionary<string, ProbeReading> ProbeReadings { get; private set; } = new Dictionary<string, ProbeReading>();
    public IReadOnlyList<string> ProbeParameterNames => ProbeReadings.Keys.ToList();
    public bool IsReadingLive { get; private set; }

    // 🔥 NEW: Queue for continuous readings (EXACTLY like your MasterServices pattern)
    public ConcurrentQueue<ProbeReadingQueueItem> CollectedReadings { get; } = new();

    public PlcProbeService()
    {
        plc = new ActUtlType64Class();
        _connectionString = ConfigurationManager.ConnectionStrings["EVMSDb"].ConnectionString!;
    }

    public bool IsConnected => isPlcConnected && isSerialConnected;  // ✅ Fixed

    // 🔥 SEPARATE CONNECTION METHODS
    public async Task<bool> ConnectSerialAsync()
    {
        try
        {
            OpenSerialPort();
            isSerialConnected = _serial?.IsOpen == true;
            return isSerialConnected;
        }
        catch
        {
            isSerialConnected = false;
            return false;
        }
    }

    public void StopAndCloseSerial()
    {
        StopLiveReading();     // cancel task
        Thread.Sleep(100);     // allow loop to exit

        try
        {
            if (_serial != null)
            {
                if (_serial.IsOpen)
                    _serial.Close();

                _serial.Dispose();
                _serial = null;
            }

            isSerialConnected = false;
        }
        catch { }
    }

    //public async Task<bool> ConnectPlcAsync()
    //{
    //    try
    //    {
    //        plc.Close();                    // Close stale
    //        plc.ActLogicalStationNumber = 1;

    //        var openTask = Task.Run(() => plc.Open());
    //        var completedTask = await Task.WhenAny(openTask, Task.Delay(3000));

    //        if (completedTask == openTask)
    //        {
    //            bool success = openTask.Result == 0;
    //            isPlcConnected = true;
    //            return success;
    //        }
    //        plc.Close();
    //        isPlcConnected = false;
    //        return false;
    //    }
    //    catch
    //    {
    //        plc.Close();
    //        isPlcConnected = false;
    //        return false;
    //    }
    //}

    // 🔥 FULL CONNECT (both)
    public async Task<bool> ConnectAsync()
    {
        //bool plcOk = await ConnectPlcAsync();
        bool serialOk = await ConnectSerialAsync();
        return serialOk;
    }

    // 🔥 SEPARATE DISCONNECT
    public void DisconnectSerial()
    {
        try
        {
            _serial?.Close();
            _serial?.Dispose();
            _serial = null;
            isSerialConnected = false;
        }
        catch { }
    }

    public void DisconnectPlc()
    {
        try
        {
            plc.Close();
            isPlcConnected = false;
        }
        catch { }
    }

    public void Disconnect()
    {
        DisconnectPlc();
        DisconnectSerial();
    }

    // ---------- PUBLIC API ----------
    public async Task LoadProbesAsync(string partNo)
    {
        var probes = new Dictionary<string, ProbeReading>();

        using var con = new SqlConnection(_connectionString);
        await con.OpenAsync();

        string query = @"
        SELECT ProbeName, ParameterName, BoxId, ChannelId 
        FROM ProbeInstallationData 
        WHERE PartNo = @P
        ORDER BY ProbeName, ChannelId";

        using var cmd = new SqlCommand(query, con);
        cmd.Parameters.AddWithValue("@P", partNo);

        using var reader = await cmd.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            string probeName = reader["ProbeName"]?.ToString() ?? "";
            string paramName = reader["ParameterName"]?.ToString() ?? "";
            int boxId = Convert.ToInt32(reader["BoxId"]);
            int channel = Convert.ToInt32(reader["ChannelId"]);

            string key = probeName;  // ✅ Unique key

            probes[key] = new ProbeReading
            {
                ParameterName = paramName,
                ProbeName = probeName,
                BoxId = boxId,
                Channel = channel,
                Value = 0,
                Status = "Not reading",
                InRange = false
            };
        }

        ProbeReadings = probes;
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

    // 🔥 LIVE READING
    public void StartLiveReading(int intervalMs = 15)
    {
        if (!isSerialConnected)
            throw new InvalidOperationException("Connect serial first.");

        if (ProbeReadings.Count == 0)
            throw new InvalidOperationException("Load probes first.");

        OpenSerialPort();  // Ensure open

        StopLiveReading();
        while (CollectedReadings.TryDequeue(out _)) { }

        _cts = new CancellationTokenSource();
        var token = _cts.Token;
        _readTask = Task.Run(() => SerialLoopAsync(token, intervalMs));
        IsReadingLive = true;
    }

    public void StopLiveReading()
    {
        _cts?.Cancel();
        if (_readTask != null)
            _ = Task.WhenAny(_readTask, Task.Delay(500));
        _cts?.Dispose();
        _cts = null;
        _readTask = null;
        IsReadingLive = false;
    }

    // 🔥 EXISTING API
    public ProbeReading? GetProbeReading(string parameterName)
    {
        ProbeReadings.TryGetValue(parameterName, out var reading);
        return reading;
    }

    public double? GetProbeValue(string parameterName)
    {
        if (ProbeReadings.TryGetValue(parameterName, out var reading))
            return reading.Value;
        return null;
    }

    // ---------- PRIVATE IMPLEMENTATION ----------
    private void OpenSerialPort()
    {
        try
        {
            _serial?.Close();
            _serial?.Dispose();
            _serial = null;

            var config = GetSerialPortConfig().FirstOrDefault();
            if (config != null && !string.IsNullOrEmpty(config.ComPort))
            {
                string comPort = config.ComPort;
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

    private async Task SerialLoopAsync(CancellationToken token, int intervalMs)
    {
        try
        {
            var stopwatch = new System.Diagnostics.Stopwatch();
            while (!token.IsCancellationRequested)
            {
                stopwatch.Restart();

                var boxes = ProbeReadings.Values.Select(p => p.BoxId).Distinct().Where(b => b > 0).ToList();
                foreach (int box in boxes)
                {
                    token.ThrowIfCancellationRequested();
                    await ReadBoxOnceAsync(box, token);

                    foreach (var kvp in ProbeReadings.Where(k => k.Value.BoxId == box))
                    {
                        var reading = kvp.Value;
                        CollectedReadings.Enqueue(new ProbeReadingQueueItem
                        {
                            ParameterName = reading.ProbeName,  // "HD001"
                            Value = reading.Value,
                            Timestamp = DateTime.Now
                        });
                    }
                }

                long elapsed = stopwatch.ElapsedMilliseconds;
                int remaining = (int)(intervalMs - elapsed);
                if (remaining > 0)
                    await Task.Delay(remaining, token);
            }
        }
        catch (OperationCanceledException) { }
    }

    private async Task ReadBoxOnceAsync(int box, CancellationToken token)
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
            MarkBoxError(box, "ERR");
            return;
        }

        string resp = await Task.Run(() => ReadFullResponse(80, token), token);
        if (string.IsNullOrEmpty(resp))
        {
            MarkBoxError(box, "ERR");
            return;
        }

        double[] values = ParseResponse(resp);
        UpdateProbesFromBox(box, values);
    }

    private void MarkBoxError(int box, string status)
    {
        foreach (var kvp in ProbeReadings.Where(k => k.Value.BoxId == box).ToList())
            UpdateProbeReading(kvp.Key, 0, status, false);
    }

    private void UpdateProbesFromBox(int box, double[] values)
    {
        foreach (var kvp in ProbeReadings.Where(k => k.Value.BoxId == box))
        {
            var reading = kvp.Value;
            if (reading.Channel < 1 || reading.Channel > 4)
            {
                UpdateProbeReading(kvp.Key, 0, "CH ERR", false);
                continue;
            }

            double value = values[reading.Channel];
            if (double.IsNaN(value))
            {
                UpdateProbeReading(kvp.Key, 0, "ERR", false);
                continue;
            }

            bool ok = value >= 0.2 && value <= 1.2;
            UpdateProbeReading(kvp.Key, value, $"{value:0.000}", ok);
        }
    }

    private void UpdateProbeReading(string paramName, double value, string status, bool inRange)
    {
        if (ProbeReadings.TryGetValue(paramName, out var reading))
        {
            reading.Value = value;
            reading.Status = status;
            reading.InRange = inRange;
        }
    }

    private string ReadFullResponse(int timeoutMs, CancellationToken token)
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

    public void Dispose()
    {
        StopLiveReading();
        Disconnect();
        try
        {
            plc?.Close();
            _serial?.Close();
            _serial?.Dispose();
        }
        catch { }
    }
}

// 🔥 DATA MODELS (unchanged)
public class ProbeReading
{
    public string ParameterName { get; set; } = "";
    public int BoxId { get; set; }
    public int Channel { get; set; }
    public double Value { get; set; }
    public string Status { get; set; } = "";
    public bool InRange { get; set; }
    public string ProbeName { get; set; } = "";
}

public class SerialPortConfigModel
{
    public int ID { get; set; }
    public string? ComPort { get; set; }
    public int BaudRate { get; set; }
}

public class ProbeReadingQueueItem
{
    public string ParameterName { get; set; } = "";
    public double Value { get; set; }
    public DateTime Timestamp { get; set; }
}
