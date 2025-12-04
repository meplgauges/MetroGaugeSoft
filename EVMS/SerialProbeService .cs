using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO.Ports;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace MetroGaugeSoft
{
    public class SerialProbeService : IDisposable
    {
        private SerialPort? _serial;
        private CancellationTokenSource? _cts;

        private readonly int[] _pos = { 4, 16, 28, 40 };

        public bool IsConnected => _serial?.IsOpen ?? false;
        public event Action<int, double[]>? BoxValuesReceived;

        public void Connect(string portName, int baud = 115200)
        {
            Disconnect();

            _serial = new SerialPort(portName, baud, Parity.None, 8, StopBits.One)
            {
                ReadTimeout = 400,
                WriteTimeout = 400
            };
            _serial.Open();
        }

        public void Disconnect()
        {
            Stop();
            _serial?.Close();
            _serial?.Dispose();
            _serial = null;
        }

        public void Start()
        {
            if (_cts != null) return;
            _cts = new CancellationTokenSource();
            _ = Task.Run(() => LiveLoop(_cts.Token));
        }

        public void Stop()
        {
            _cts?.Cancel();
            _cts?.Dispose();
            _cts = null;
        }

        private async Task LiveLoop(CancellationToken token)
        {
            try
            {
                while (!token.IsCancellationRequested)
                {
                    await ReadAllBoxesOnce();
                    await Task.Delay(8, token);
                }
            }
            catch { }
        }

        private async Task ReadAllBoxesOnce()
        {
            if (!IsConnected) return;

            // Boxes 1–32 supported
            for (int box = 1; box <= 32; box++)
            {
                string cmd = $"*{box:D3}VALL#\r";

                try
                {
                    _serial!.DiscardInBuffer();
                    _serial.Write(cmd);
                }
                catch
                {
                    BoxValuesReceived?.Invoke(box, new double[] { double.NaN, double.NaN, double.NaN, double.NaN });
                    continue;
                }

                string resp = ReadResponse(60);
                double[] values = ParseResponse(resp);

                BoxValuesReceived?.Invoke(box, values);
            }

            await Task.CompletedTask;
        }

        private string ReadResponse(int timeout)
        {
            if (_serial == null) return "";
            string text = "";
            var sw = System.Diagnostics.Stopwatch.StartNew();

            while (sw.ElapsedMilliseconds < timeout)
            {
                text += _serial.ReadExisting();
                if (text.Contains("#")) break;
                Thread.Sleep(3);
            }

            return text;
        }

        private double[] ParseResponse(string r)
        {
            double[] ch = Enumerable.Repeat(double.NaN, 4).ToArray();

            if (string.IsNullOrWhiteSpace(r)) return ch;

            int idx = r.IndexOf("VALL", StringComparison.OrdinalIgnoreCase);
            if (idx >= 0) r = r[(idx + 4)..];

            r = r.Replace("#", "").Trim();

            if (r.Length < _pos.Last()) return ch;

            const string pattern = @"[-+]?(?:\d*\.\d+|\d+\.?\d*)";

            for (int i = 0; i < 4; i++)
            {
                int start = _pos[i];
                int len = (i == 3 ? r.Length - start : _pos[i + 1] - start);

                if (start >= r.Length || len <= 0) continue;

                string seg = r.Substring(start, len).Trim();

                var m = Regex.Match(seg, pattern);
                if (m.Success && double.TryParse(m.Value, NumberStyles.Any, CultureInfo.InvariantCulture, out double v))
                    ch[i] = v;
            }

            return ch;
        }

        public void Dispose() => Disconnect();
    }
}