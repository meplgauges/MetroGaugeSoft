using ActUtlType64Lib;
using Microsoft.Data.SqlClient;
using System.Configuration;
using System.Diagnostics;
using System.Globalization;
using System.IO.Ports;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

public sealed class PlcProbeService : IDisposable
{
    private IActUtlType64 plc;
    private SerialPort? _serial;

    private CancellationTokenSource? _cts;
    private Task? _readTask;

    private bool _isSerialConnected;
    private readonly string _connectionString;

    private int _sampleIntervalMs;
    private long _lastSampleTick;

    private Dictionary<int, List<ProbeReading>> _probesByBox = new();

    private static readonly Regex ValueRegex =
        new Regex(@"C\d{2}([-+]?\d*\.?\d+)", RegexOptions.Compiled);

    public bool IsConnected => _isSerialConnected && _serial?.IsOpen == true;
    public bool IsReadingLive { get; private set; }

    public IReadOnlyDictionary<string, ProbeReading> ProbeReadings { get; private set; }
        = new Dictionary<string, ProbeReading>();

    public PlcProbeService()
    {
        plc = new ActUtlType64Class();
        _connectionString = ConfigurationManager.ConnectionStrings["EVMSDb"].ConnectionString!;
    }

    // ================= SERIAL CONNECTION =================
    public async Task<bool> ConnectSerialAsync()
    {
        try
        {
            OpenSerialPort();
            _isSerialConnected = _serial?.IsOpen == true;
            return _isSerialConnected;
        }
        catch
        {
            _isSerialConnected = false;
            return false;
        }
    }

    private void OpenSerialPort()
    {
        _serial?.Dispose();

        var config = GetSerialPortConfig().FirstOrDefault()
            ?? throw new InvalidOperationException("Serial port config missing");

        _serial = new SerialPort(config.ComPort!, config.BaudRate, Parity.None, 8, StopBits.One)
        {
            ReadTimeout = 50,
            WriteTimeout = 50
        };

        _serial.Open();
    }

    public void StopAndCloseSerial()
    {
        StopLiveReading();

        try
        {
            _serial?.Close();
            _serial?.Dispose();
        }
        catch { }

        _serial = null;
        _isSerialConnected = false;
    }

    // ================= LOAD PROBES =================
    public async Task LoadProbesAsync(string partNo)
    {
        var probes = new Dictionary<string, ProbeReading>();

        using var con = new SqlConnection(_connectionString);
        await con.OpenAsync();

        var cmd = new SqlCommand(@"
            SELECT ProbeName, ParameterName, BoxId, ChannelId
            FROM ProbeInstallationData
            WHERE PartNo = @P
            ORDER BY ProbeName, ChannelId", con);

        cmd.Parameters.AddWithValue("@P", partNo);

        using var r = await cmd.ExecuteReaderAsync();
        while (await r.ReadAsync())
        {
            string probeName = r["ProbeName"].ToString()!;
            probes[probeName] = new ProbeReading
            {
                ProbeName = probeName,
                ParameterName = r["ParameterName"].ToString()!,
                BoxId = Convert.ToInt32(r["BoxId"]),
                Channel = Convert.ToInt32(r["ChannelId"]),
                Status = "Idle"
            };
        }

        ProbeReadings = probes;

        _probesByBox = ProbeReadings.Values
            .Where(p => p.BoxId > 0)
            .GroupBy(p => p.BoxId)
            .ToDictionary(g => g.Key, g => g.ToList());
    }

    // ================= START / STOP READING =================
    public void StartLiveReading(int intervalMs, int sampleIntervalMs)
    {
        if (!IsConnected)
            throw new InvalidOperationException("Serial not connected");

        StopLiveReading();

        _sampleIntervalMs = sampleIntervalMs;
        _lastSampleTick = 0;

        _cts = new CancellationTokenSource();
        _readTask = Task.Run(() => SerialLoopAsync(_cts.Token, intervalMs));
        IsReadingLive = true;
    }

    public void StopLiveReading()
    {
        _cts?.Cancel();
        IsReadingLive = false;
    }

    // ================= MAIN SERIAL LOOP =================
    private async Task SerialLoopAsync(CancellationToken token, int intervalMs)
    {
        var boxes = _probesByBox.Keys.ToArray();
        int boxIndex = 0;

        try
        {
            while (!token.IsCancellationRequested)
            {
                int box = boxes[boxIndex];
                boxIndex = (boxIndex + 1) % boxes.Length;

                await ReadBoxOnceAsync(box, token);

                // Optional: update a timestamp for each probe
                long now = Environment.TickCount64;
                if (now - _lastSampleTick >= _sampleIntervalMs)
                {
                    _lastSampleTick = now;
                    foreach (var p in _probesByBox[box])
                    {
                        p.LastUpdatedUtc = DateTime.UtcNow;
                    }
                }

                await Task.Delay(intervalMs, token);
            }
        }
        catch (OperationCanceledException) { }
    }

    // ================= READ SINGLE BOX =================
    private async Task ReadBoxOnceAsync(int box, CancellationToken token)
    {
        if (_serial == null || !_serial.IsOpen)
            return;

        try
        {
            _serial.DiscardInBuffer();
            _serial.Write($"*{box:D3}VALL#\r");
        }
        catch
        {
            MarkBoxError(box);
            return;
        }

        string resp = await ReadResponseAsync(70, token);
        if (string.IsNullOrWhiteSpace(resp))
        {
            MarkBoxError(box);
            return;
        }

        UpdateProbesFromBox(box, ParseResponse(resp));
    }

    private void MarkBoxError(int box)
    {
        foreach (var p in _probesByBox[box])
        {
            p.Value = 0;
            p.Status = "ERR";
            p.InRange = false;
            p.LastUpdatedUtc = DateTime.UtcNow;
        }
    }

    private void UpdateProbesFromBox(int box, double[] values)
    {
        foreach (var p in _probesByBox[box])
        {
            if (p.Channel < 1 || p.Channel > 4)
            {
                p.Status = "CH ERR";
                continue;
            }

            double v = values[p.Channel];
            if (double.IsNaN(v))
            {
                p.Status = "ERR";
                continue;
            }

            p.Value = v;
            p.Status = v.ToString("0.000", CultureInfo.InvariantCulture);
            p.InRange = v >= 0.2 && v <= 1.2;
            p.LastUpdatedUtc = DateTime.UtcNow;
        }
    }

    // ================= SERIAL RESPONSE =================
    private async Task<string> ReadResponseAsync(int timeoutMs, CancellationToken token)
    {
        var sw = Stopwatch.StartNew();
        var sb = new StringBuilder();

        while (sw.ElapsedMilliseconds < timeoutMs)
        {
            token.ThrowIfCancellationRequested();

            if (_serial!.BytesToRead > 0)
            {
                sb.Append(_serial.ReadExisting());
                if (sb.ToString().Contains('#'))
                    break;
            }

            await Task.Yield(); // ✅ REQUIRED
        }

        return sb.ToString();
    }



    private double[] ParseResponse(string r)
    {
        var ch = new double[5];
        for (int i = 1; i <= 4; i++) ch[i] = double.NaN;

        var matches = ValueRegex.Matches(r);
        for (int i = 0; i < matches.Count && i < 4; i++)
        {
            double.TryParse(matches[i].Groups[1].Value,
                NumberStyles.Any,
                CultureInfo.InvariantCulture,
                out ch[i + 1]);
        }
        return ch;
    }

    // ================= DB =================
    private List<SerialPortConfigModel> GetSerialPortConfig()
    {
        var list = new List<SerialPortConfigModel>();

        using var con = new SqlConnection(_connectionString);
        using var cmd = new SqlCommand("SELECT * FROM SerialSettings WHERE ID = 1", con);

        con.Open();
        using var r = cmd.ExecuteReader();
        while (r.Read())
        {
            list.Add(new SerialPortConfigModel
            {
                ID = (int)r["ID"],
                ComPort = r["ComPort"].ToString(),
                BaudRate = (int)r["BaudRate"]
            });
        }
        return list;
    }

    public void Dispose()
    {
        StopLiveReading();
        StopAndCloseSerial();
        try { plc?.Close(); } catch { }
    }
}

// ================= MODELS =================
public class ProbeReading
{
    public string ProbeName { get; set; } = string.Empty;
    public string ParameterName { get; set; } = string.Empty;
    public int BoxId { get; set; }
    public int Channel { get; set; }
    public double Value { get; set; }
    public string Status { get; set; } = string.Empty;
    public bool InRange { get; set; }

    public DateTime LastUpdatedUtc { get; set; } // optional timestamp
}

public class SerialPortConfigModel
{
    public int ID { get; set; }
    public string? ComPort { get; set; }
    public int BaudRate { get; set; }
}
