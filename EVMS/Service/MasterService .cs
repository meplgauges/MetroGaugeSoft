using ActUtlType64Lib;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Concurrent;
using System.Diagnostics;
using System.Net.Sockets;
using System.Text;
using System.Windows;
using System.Windows.Input.Manipulations;
using System.Windows.Navigation;
using Windows.Media.Protection.PlayReady;
using WinRT;

namespace EVMS.Service
{
    public class ProbeMeasurement
    {
        public string ProbeId { get; set; } = "";      // Unique key: e.g. Diameter_CH1
        public string Name { get; set; } = "";         // Logical parameter name: Diameter
        public List<double> Readings { get; set; } = new List<double>();
        public double MaxValue { get; set; } = 0;
        public double MinValue { get; set; } = 0;
        public double MasterValue { get; set; } = 0;
        public double TolerancePlus { get; set; } = 0;
        public double ToleranceMinus { get; set; } = 0;
        public int SignChange { get; set; } = 0;
        public double Compensation { get; set; } = 0;
    }



    public class ParameterResult
    {
        public double Min { get; set; }
        public double Value { get; set; }
        public bool IsOk { get; set; }
    }


    public class MasterService : IDisposable
    {
        public event Action? MeasurementStarted;
        public event Action? MeasurementStopped;

        public bool _continueMeasurement = false;
        public bool _continueMastring = false;
        private bool _isCameraReady = false;

        public bool IsMeasurementRunning { get; set; } = false;
        private TcpClient _client;
        private NetworkStream _stream;



        public event Action<string>? StatusMessageUpdated;
        private List<ProbeMeasurement> _orderedProbeMeasurements = new List<ProbeMeasurement>();

        public delegate void CalculatedValuesWithStatusHandler(object? sender, Dictionary<string, ParameterResult> results);
        public event CalculatedValuesWithStatusHandler? CalculatedValuesWithStatusReady;
        public event EventHandler<MasterCompletedEventArgs>? MasterCompleted;
        public delegate void CalculatedValuesReadyHandler(object? sender, Dictionary<string, double> calculatedValues);
        public event CalculatedValuesReadyHandler? CalculatedValuesReady;
        private bool _isMeasurementRunning = true;


        private IActUtlType64 plc;
        private readonly DataStorageService _dataStorageService;
        private readonly PlcProbeService _plcProbeService;

        private Dictionary<string, ProbeMeasurement> _probeMeasurements = new Dictionary<string, ProbeMeasurement>();
        private int ArraySize;


        private readonly ConcurrentQueue<(string ProbeId, double Value)> _collectedReadings
            = new ConcurrentQueue<(string, double)>();

        private int _currentOperationalMode = 1;
       private string _currentPartCode = "";

        public string ActiveIdNo { get; private set; } // ✅ Public access for calculations

        // 🔔 UI notification
        public event Action? ResetRequested;
        public int ActiveIdValue { get; private set; } = 0;  // ✅ Public access for calculations
        public int ActiveBotValue { get; private set; } = 0;  // ✅ Public access for calculations
        public bool IsMasteringStage { get; set; } = true;
        public bool MasterComplete { get; set; } = false;
        public bool MasteringOK { get; set; } = false;
        public bool Abort { get; set; } = false;
        public double MinReferenceValue { get; set; } = 0;


        //private const string MotorOnDevice = "M14";
        //private const string Auto  = "X14";

        public MasterService()
        {
            //var resultPage = new ResultPage();

            _dataStorageService = new DataStorageService();
            _plcProbeService = new PlcProbeService();
            plc = new ActUtlType64Class { ActLogicalStationNumber = 1 };
            ArraySize = _dataStorageService.GetReadingCount();
            //SetPlcDevice("M1", 1); //Software Ready
            //LoadProbeConfigurationsforMasterInspection(_currentPartCode);
            // ApplyActiveIdPlcBits();
        }
        public async Task StartMeasurementAsync()
        {
            _continueMeasurement = true;
            await RunMeasurementCycleAsync();
        }

        public void StopMeasurement()
        {
            _continueMeasurement = false;
        }

        protected virtual void OnCalculatedValuesWithStatusReady(Dictionary<string, ParameterResult> results)
        {
            CalculatedValuesWithStatusReady?.Invoke(this, results);
        }
        public bool IsConnected => _plcProbeService?.IsConnected ?? false;

        // 🔹 Connect with timeout
        private async Task<TcpClient?> ConnectIPAsync(string ip, int port, int timeoutMs = 3000)
        {
            try
            {
                var client = new TcpClient();
                var connectTask = client.ConnectAsync(ip, port);
                var timeoutTask = Task.Delay(timeoutMs);

                var completedTask = await Task.WhenAny(connectTask, timeoutTask);

                if (completedTask == timeoutTask || !client.Connected)
                {
                    client.Dispose();
                    return null;
                }

                return client;
            }
            catch
            {
                return null;
            }
        }



        public async Task<bool> TestConnectionAsync(string productCode)
        {
            _isCameraReady = false;

            try
            {
                var config = _dataStorageService.GetCameraConfiguration();

                if (config == null)
                {
                    MessageBox.Show(
                        "Camera configuration not found in database.",
                        "Configuration Error",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);
                    return false;
                }

                string ip1 = config.Camera1IP;
                string ip2 = config.Camera2IP;
                int port = config.Port;

                string command = $"do productchange \"{productCode}\"\r\n";

                var connectTask1 = ConnectIPAsync(ip1, port);
                var connectTask2 = ConnectIPAsync(ip2, port);

                await Task.WhenAll(connectTask1, connectTask2);

                TcpClient? client1 = connectTask1.Result;
                TcpClient? client2 = connectTask2.Result;

                // ❌ Connection error
                if (client1 == null || client2 == null)
                {
                    MessageBox.Show(
                        $"Connection failed.\n" +
                        $"{(client1 == null ? $"Device1 ({ip1}) not reachable.\n" : "")}" +
                        $"{(client2 == null ? $"Device2 ({ip2}) not reachable." : "")}",
                        "Connection Error",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);

                    client1?.Dispose();
                    client2?.Dispose();
                    return false;
                }

                using (client1)
                using (client2)
                using (NetworkStream stream1 = client1.GetStream())
                using (NetworkStream stream2 = client2.GetStream())
                {
                    byte[] data = Encoding.ASCII.GetBytes(command);

                    await Task.WhenAll(
                        stream1.WriteAsync(data, 0, data.Length),
                        stream2.WriteAsync(data, 0, data.Length)
                    );

                    byte[] buffer1 = new byte[1024];
                    byte[] buffer2 = new byte[1024];

                    var readTask1 = stream1.ReadAsync(buffer1, 0, buffer1.Length);
                    var readTask2 = stream2.ReadAsync(buffer2, 0, buffer2.Length);

                    await Task.WhenAll(readTask1, readTask2);

                    bool device1Ok = false;
                    bool device2Ok = false;

                    if (readTask1.Result > 0)
                    {
                        string response1 = Encoding.ASCII.GetString(buffer1, 0, readTask1.Result);
                        device1Ok = response1.Contains("OK");
                    }

                    if (readTask2.Result > 0)
                    {
                        string response2 = Encoding.ASCII.GetString(buffer2, 0, readTask2.Result);
                        device2Ok = response2.Contains("OK");
                    }

                    // ❌ Command error
                    if (!device1Ok || !device2Ok)
                    {
                        MessageBox.Show(
                            $"{(!device1Ok ? $"Device1 ({ip1}) did not acknowledge.\n" : "")}" +
                            $"{(!device2Ok ? $"Device2 ({ip2}) did not acknowledge." : "")}",
                            "Command Error",
                            MessageBoxButton.OK,
                            MessageBoxImage.Warning);

                        return false;
                    }

                    // ✅ SUCCESS → NO MESSAGE BOX
                    _isCameraReady = true;
                    return true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Communication error:\n" + ex.Message,
                    "Communication Error",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);

                return false;
            }
        }

        public async Task<bool> EnsureConnectionAsync()
        {
           // await _plcProbeService.ConnectAsync();
            //LOCAL PLC (async + safe)
            try
            {
                plc.Close();  // ✅ Close stale FIRST
                plc.ActLogicalStationNumber = 1;
                var openTask = Task.Run(() => plc.Open());
                var completed = await Task.WhenAny(openTask, Task.Delay(3000));

                if (completed == openTask && openTask.Result == 0)
                {
                   // Debug.WriteLine("✅ Local PLC ready");

                    var autoList = _dataStorageService.GetActiveBit();

                    // Find the control with Code "LS"
                    var shControl = autoList.FirstOrDefault(c => c.Code == "LS");
                    var CLControl = autoList.FirstOrDefault(c => c.Code == "CL");


                    // Check if shControl exists and its value is 1
                    if (shControl != null && shControl.Bit == 1)
                    {
                        // Set PLC bit to 1
                        SetPlcDevice("M5", 1);
                    }

                    //if (CLControl != null && CLControl.Bit == 1)
                    //{


                            
                    //        SetPlcDevice("M6", 1);

                    //    await TestConnectionAsync();

                        
                    //}

                   await ProcessActivePartAndSetRoboBit();

                    return true;
                }
                else
                {
                    string err = completed == openTask ? $"Error {openTask.Result}" : "Timeout";
                    MessageBox.Show($"Local PLC failed: {err}", "PLC Init Error",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                    plc.Close();
                    return false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"PLC error: {ex.Message}", "PLC Error");
                plc.Close();
                return false;
            }
        }



        public void ResetAll()
        {
           
        ResetRequested?.Invoke(); }

        //private void ProbeValueUpdatedHandler(object? sender, ProbeReadingEventArgs e)
        //{
        //    _collectedReadings.Enqueue((e.ModuleId, e.Value));
        //}

        public async Task ProcessActivePartAndSetRoboBit()
        {
            // 1️⃣ Determine Auto/Manual mode
            var autoList = _dataStorageService.GetActiveBit();
            var autoControl = autoList?.FirstOrDefault(c =>
                string.Equals(c.Description, "Auto/Manual", StringComparison.OrdinalIgnoreCase));
            //int bitValue = GetPlcDeviceBit("X0"); // PLC Auto/Manual bit

            bool isAuto = autoControl != null && autoControl.Bit == 1;

            if (!isAuto)
            {
                return; // Manual mode - skip auto processing
            }

            // 2️⃣ Get active parts
            var activeParts = _dataStorageService.GetActiveParts();
            if (activeParts == null || activeParts.Count == 0)
            {
                MessageBox.Show("No active parts found.", "Info", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            // Take first active part
            var activePart = activeParts[0];
            string partNumber = activePart?.Para_No ?? "";
            if (string.IsNullOrEmpty(partNumber))
            {
                MessageBox.Show("Active part has no Para_No.", "Info", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            // 3️⃣ Get BOTH RoboBit and LaserBit configuration
            var bitConfig = _dataStorageService.GetBitConfigByPartNo(partNumber);
            if (!bitConfig.HasRoboBit)
            {
                MessageBox.Show($"No RoboBit found for part {partNumber}.", "Info",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }


            var CLControl = autoList.FirstOrDefault(c => c.Code == "CL");

            if (CLControl != null && CLControl.Bit == 1)
            {



                SetPlcDevice("M6", 1);

               // await TestConnectionAsync(partNumber);


            }
            try
            {
                // Set RoboBit (primary probe)
                if (bitConfig.HasRoboBit)
                {
                    SetPlcDevice(bitConfig.RoboBit, 1);
                }

                // Set LaserBit (secondary measurement)
                if (bitConfig.HasLaserBit)
                {
                    SetPlcDevice(bitConfig.LaserBit, 1);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to set bits for {partNumber}: {ex.Message}", "PLC Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }


        public bool SetPlcDevice(string device, int value)
        {
            int ret = plc.SetDevice(device, (short)value);
            if (ret != 0)
            {
                //NotifyStatus("SetDevice failed!!");

                return false;
            }
            return true;
        }

        public int GetPlcDeviceBit(string device)
        {
            if (plc.GetDevice(device, out int value) != 0)
            {
                //NotifyStatus($"GetDevice failed for bit {device}! Exception");
                return -1; // Indicate failure
            }
            return value; // Return the bit value read from PLC device
        }


        private void Cleanup()
        {
            _plcProbeService.StopLiveReading();  // ✅ ADD
            _plcProbeService.Dispose();          // ✅ ADD
            plc?.Close();
        }

        public enum ProcedureMode
        {
            Mastering,
            MasterInspection,
            Measurement
        }




        private async Task WaitForPlcBitAsync(string bitName, int timeoutMs = 10000)
        {
            int elapsed = 0;

            while (GetPlcDeviceBit(bitName) != 1)
            {
                await Task.Delay(100);
                elapsed += 50;

                if (elapsed >= timeoutMs)
                {
                    MessageBox.Show("PLC Error: {bitName} did not turn ON.");
                    SetPlcDevice("L25", 1);

                }
            }
        }



        public async Task MasterCheckProcedureAsync(ProcedureMode mode)
        {
            try
            {
                var autoList = _dataStorageService.GetActiveBit();

                var activeIdPart = _dataStorageService
                    .GetActiveID()
                    .FirstOrDefault(p => p.BOT_Value >= 0);

                var CLControl = autoList.FirstOrDefault(c => c.Code == "CL");

                ActiveIdNo = activeIdPart?.Para_No;

                ActiveIdValue = activeIdPart?.ID_Value ?? 0;
                ActiveBotValue = activeIdPart?.BOT_Value ?? 0;

                
                if (autoList == null || autoList.Count == 0)
                {
                    MessageBox.Show("No Settings Found.", "Info", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                var autoControl = autoList.FirstOrDefault(c => string.Equals(c.Description, "Auto/Manual", StringComparison.OrdinalIgnoreCase));
                int bitValue = GetPlcDeviceBit("X0");

                if ((autoControl?.Bit == 1 && bitValue == 1) || (autoControl?.Bit == 0 && bitValue == 0))
                {
                    // Modes matched: proceed with the rest of the procedure
                }
                else
                {
                    if (autoControl?.Bit == 0 && bitValue == 1)
                    {
                        MessageBox.Show("Software is in Manual mode. Please switch PLC to Manual mode.", "Info", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    else if (autoControl?.Bit == 1 && bitValue == 0)
                    {
                        MessageBox.Show("Software is in Auto mode. Change PLC to Auto mode.", "Info", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    return;  // Exit early if modes do not match
                }


                var activeParts = _dataStorageService.GetActiveParts();
                if (activeParts == null || activeParts.Count == 0)
                {
                    MessageBox.Show("No active parts found.", "Info", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                _currentPartCode = activeParts[0]?.Para_No ?? "";
                if (string.IsNullOrEmpty(_currentPartCode))
                {
                    MessageBox.Show("Active part code is invalid.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }


                switch (ActiveBotValue)
                {
                    case 1:
                        SetPlcDevice("M129", 0);
                        SetPlcDevice("M130", 0);
                        SetPlcDevice("M128", 1);
                        break;

                    case 2:
                        SetPlcDevice("M130", 0);
                        SetPlcDevice("M128", 0);
                        SetPlcDevice("M129", 1);
                        break;

                    case 3:
                        SetPlcDevice("M128", 0);
                        SetPlcDevice("M129", 0);
                        SetPlcDevice("M130", 1);
                        break;

                    case 0:
                    default:
                        // Do nothing
                        break;
                }


                if (autoControl?.Bit == 1)
                {
                    SetPlcDevice("M1", 1); //Software Ready
                    SetPlcDevice("M16", 0);
                    SetPlcDevice("M26", 0);
                    SetPlcDevice("M11", 0);


                    if (CLControl != null && CLControl.Bit == 1)
                    {
                        // CL is ON → Camera check required WHEN THE CAMERA INTERLOCK IS ON

                        if (!_isCameraReady)
                        {
                            MessageBox.Show("Camera is not ready. Operation blocked.",
                                            "Camera Error",
                                            MessageBoxButton.OK,
                                            MessageBoxImage.Warning);
                            return;   
                        }
                    }


                    if (GetPlcDeviceBit("M3") != 0)
                    {
                        MessageBox.Show(
                            "Previous cycle is not cleared.\n" +
                            "Please clear the cycle before starting.",
                            "Auto Mode",
                            MessageBoxButton.OK,
                            MessageBoxImage.Warning);

                        return;
                    }
                    switch (mode)
                    {
                        case ProcedureMode.Mastering:
                            SetPlcDevice("M16", 1);
                            SetPlcDevice("M26", 0);//Measurment off
                            SetPlcDevice("M11", 0);//ClearCycle off
                            break;
                        case ProcedureMode.MasterInspection:
                            SetPlcDevice("M16", 1);
                            break;
                        case ProcedureMode.Measurement:
                            SetPlcDevice("M26", 1);
                            LoadProbeConfigurationsforMasterInspection(_currentPartCode);

                            break;
                    }
                }




                if (mode == ProcedureMode.MasterInspection)
                {
                    LoadProbeConfigurations(_currentPartCode);

                }
                else
                {
                    LoadProbeConfigurationsforMasterInspection(_currentPartCode);

                }
                var sortedProbeMeasurements = _orderedProbeMeasurements;

                foreach (var pm in sortedProbeMeasurements)
                {
                    pm.Readings.Clear();
                    pm.MaxValue = 0;
                    pm.MinValue = 0;
                }

                if (mode == ProcedureMode.Measurement)
                {
                    // Start the measurement cycle asynchronously her
                    return;
                }


                bool firstMeasurementCycle = true;
                if (mode == ProcedureMode.Mastering || mode == ProcedureMode.MasterInspection)
                {
                    if (autoControl?.Bit == 0)
                    { 

                        string loadMsg = mode == ProcedureMode.Mastering ? "LOAD THE VALUE IN FIXTURE..." : " PLACE PART FOR MEASUREMENT...";
                        await NotifyOnUIAsync(loadMsg);
                        while (GetPlcDeviceBit("X46") != 1) await Task.Delay(100);
                        int Check = GetPlcDeviceBit("M4");
                        if (Check != 1)
                        {
                            MessageBox.Show("Some Cylinders not at home.");
                            SetPlcDevice("L25", 1);// Home Bit For ForceFully Given 
                            return;
                        }
                        string promptMsg = mode == ProcedureMode.Mastering ? "PRESS START SWITCH TO START MASTERING" : "PRESS START SWITCH TO START MasterInspection";
                        await NotifyOnUIAsync(promptMsg);
                        while (GetPlcDeviceBit("X1") != 1) await Task.Delay(100);
                        await NotifyOnUIAsync("START BUTTON PRESSED");


                        await NotifyOnUIAsync("MOTOR BLOCK CYCLE ON");
                        await WaitForPlcBitAsync("X51");
                        SetPlcDevice("M114", 1);

                        await NotifyOnUIAsync("MOTOR DOWN SIGNAL ON");
                        await WaitForPlcBitAsync("X50");
                        SetPlcDevice("M127", 1);

                        switch (ActiveIdValue)
                        {

                            case 1:
                                await NotifyOnUIAsync("RIGHT_ID_CYL 9.508 AND LEFT_ID_CYL 9.508 ON ");

                                await WaitForPlcBitAsync("X65"); //Right_Block_ID-12.642
                                await WaitForPlcBitAsync("X75"); //Right_Block_ID-8.072
                                await WaitForPlcBitAsync("X71"); //LEFT_BLOCK_ID-12.642
                                await WaitForPlcBitAsync("X101"); //LEFT_BLOCK_ID-8.072
                                SetPlcDevice("M131", 1);//Right_Block_ID-9.508
                                SetPlcDevice("M132", 1);//LEFT_BLOCK_ID-9.508
                                break;
                            case 2:

                                await NotifyOnUIAsync("RIGHT_BLOCK_CYL 12.642 AND LEFT_BLOCK_CYL 12.642 ON");

                                await WaitForPlcBitAsync("X67");
                                await WaitForPlcBitAsync("X61");//Right_Block_ID-9.508
                                await WaitForPlcBitAsync("X75"); //Right_Block_ID-8.072
                                await WaitForPlcBitAsync("X101");//LEFT_BLOCK_ID-8.072
                                await WaitForPlcBitAsync("X63");//LEFT_BLOCK_ID-9.508
                                await WaitForPlcBitAsync("X73");//LEFT_BLOCK_ID-12.642

                                SetPlcDevice("M116", 1);//Right Cylinder ID-12.642
                                SetPlcDevice("M118", 1);//lEFT Cylinder ID-12.642

                                await NotifyOnUIAsync("RIGHT_ID 12.642 AND LEFT_ID 12.642 ON");

                                await WaitForPlcBitAsync("X66");
                                await WaitForPlcBitAsync("X72");

                                SetPlcDevice("M133", 1);//Right Cylinder ID-12.642
                                SetPlcDevice("M134", 1);//lEFT Cylinder ID-12.642

                                break;

                            case 3:
                                await NotifyOnUIAsync("RIGHT_BLOCK_CYL 8.072 AND LEFT_BLOCK_CYL 8.072 ON");

                                await WaitForPlcBitAsync("X77");//Right_CYC_ID-8.072
                                await WaitForPlcBitAsync("X61");//Right_Block_ID-9.508
                                await WaitForPlcBitAsync("X65");//Right_Block_ID-12.642
                                await WaitForPlcBitAsync("X103");//LEFT_CYC_ID-8.072
                                await WaitForPlcBitAsync("X73");//LEFT_BLOCK_ID-12.642
                                await WaitForPlcBitAsync("X63");//LEFT_BLOCK_ID-9.508


                                SetPlcDevice("M120", 1);//Right_Block_ID-8.072 ON
                                SetPlcDevice("M122", 1);//LEFT_BLOCK_ID-8.072

                                await NotifyOnUIAsync("RIGHT_ID 8.072 AND LEFT_ID  8.072 ON");

                                await WaitForPlcBitAsync("X76"); //Right_CYC_ID - 8.072 
                                await WaitForPlcBitAsync("X102");//LEFT_CYC_ID-8.072

                                SetPlcDevice("M135", 1);//Right_ID-8.072 ON
                                SetPlcDevice("M136", 1);//LEFT_ID-8.072 ON

                                break;

                            case 0:
                                // Do nothing
                                break;

                            default:
                                // Optional: handle unexpected values
                                break;
                        }


                    }
                    else
                    {
                        // ================= AUTO MODE PRE-CHECK =================

                        // 1️⃣ Check cylinders at home
                        if (GetPlcDeviceBit("M4") != 1)
                        {
                            MessageBox.Show(
                                "Some cylinders are not at home.\nForcing homing operation.",
                                "Auto Mode",
                                MessageBoxButton.OK,
                                MessageBoxImage.Warning);

                            SetPlcDevice("L25", 1); // Force Home
                            return;
                        }

                        // 2️⃣ Check previous cycle cleared


                        // ✅ All checks passed → Auto mode can continue


                        string startMsg = mode == ProcedureMode.Mastering ? "PRESS THE ROBO START BUTTON TO BEGIN MASTERING.  \r\n" : "PRESS THE ROBO START BUTTON TO BEGIN MASTER INSPECTION";
                        await NotifyOnUIAsync(startMsg);
                        while (GetPlcDeviceBit("B2") != 1) await Task.Delay(100);
                        //await NotifyOnUIAsync("Robo start button Pressed");//DIRECTLY GETTING THE ROBO START

                        //SetPlcDevice("M101", 1); //Robo Start Bit
                        //if (!_continueMastring) return;

                        await NotifyOnUIAsync("WAITING FOR ROBOT TO LOAD PART AND REACH SAFE POSITION...");

                        while (GetPlcDeviceBit("M18") != 1) await Task.Delay(100);

                        //if (_continueMeasurement)
                      while (GetPlcDeviceBit("X46") != 1) await Task.Delay(100);


                        await NotifyOnUIAsync("MOTOR BLOCK CYCLE ON");
                        await WaitForPlcBitAsync("X51");
                        SetPlcDevice("M114", 1);

                        await NotifyOnUIAsync("MOTOR DOWN ");
                        await WaitForPlcBitAsync("X50");
                        SetPlcDevice("M127", 1);

                        switch (ActiveIdValue)
                        {

                            case 1:
                                await NotifyOnUIAsync("RIGHT_ID_CYL 9.508 AND LEFT_ID_CYL 9.508 ON ");

                                await WaitForPlcBitAsync("X65"); //Right_Block_ID-12.642
                                await WaitForPlcBitAsync("X75"); //Right_Block_ID-8.072
                                await WaitForPlcBitAsync("X71"); //LEFT_BLOCK_ID-12.642
                                await WaitForPlcBitAsync("X101"); //LEFT_BLOCK_ID-8.072
                                SetPlcDevice("M131", 1);//Right_Block_ID-9.508
                                SetPlcDevice("M132", 1);//LEFT_BLOCK_ID-9.508
                                break;
                            case 2:

                                await NotifyOnUIAsync("RIGHT_BLOCK_CYL 12.642 AND LEFT_BLOCK_CYL 12.642 ON");

                                await WaitForPlcBitAsync("X67");
                                await WaitForPlcBitAsync("X61");
                                await WaitForPlcBitAsync("X75");
                                await WaitForPlcBitAsync("X101");
                                await WaitForPlcBitAsync("X63");
                                await WaitForPlcBitAsync("X73");

                                SetPlcDevice("M116", 1);//Right Cylinder ID-12.642
                                SetPlcDevice("M118", 1);//lEFT Cylinder ID-12.642

                                // Wait sequence – must complete before setting M
                                await NotifyOnUIAsync("RIGHT_ID 12.642 AND LEFT_ID 12.642 ON");

                                await WaitForPlcBitAsync("X66");
                                await WaitForPlcBitAsync("X72");

                                SetPlcDevice("M133", 1);//Right Cylinder ID-12.642
                                SetPlcDevice("M134", 1);//lEFT Cylinder ID-12.642

                                break;

                            case 3:
                                //CHECK ALL HAVE TO HOME POSITION
                                await NotifyOnUIAsync("RIGHT_BLOCK_CYL 8.072 AND LEFT_BLOCK_CYL 8.072 ON");

                                await WaitForPlcBitAsync("X77");//Right_CYC_ID-8.072
                                await WaitForPlcBitAsync("X61");//Right_Block_ID-9.508
                                await WaitForPlcBitAsync("X65");//Right_Block_ID-12.642
                                await WaitForPlcBitAsync("X103");//LEFT_CYC_ID-8.072
                                await WaitForPlcBitAsync("X73");//LEFT_BLOCK_ID-12.642
                                await WaitForPlcBitAsync("X63");//LEFT_BLOCK_ID-9.508


                                SetPlcDevice("M120", 1);//Right_Block_ID-8.072 ON
                                SetPlcDevice("M122", 1);//LEFT_BLOCK_ID-8.072

                                //// Wait sequence – must complete before setting M
                                await NotifyOnUIAsync("RIGHT_ID 8.072 AND LEFT_ID 8.072 ON");

                                await WaitForPlcBitAsync("X76"); //Right_CYC_ID - 8.072 
                                await WaitForPlcBitAsync("X102");//LEFT_CYC_ID-8.072

                                SetPlcDevice("M135", 1);//Right_ID-8.072 ON
                                SetPlcDevice("M136", 1);//LEFT_ID-8.072 ON

                                break;

                            case 0:
                                // Do nothing
                                break;


                            default:
                                // Optional: handle unexpected values
                                break;
                        }
                    }
       

                    await RunMotorAndCollectReadingsAsync(sortedProbeMeasurements, mode);

             

                    await HandleProcedurePostProcessingAsync(mode, sortedProbeMeasurements, firstMeasurementCycle);
                    
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error in MasterCheckProcedure: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // Helper should continue method

        public async Task RunMeasurementCycleAsync()
        {
            var autoList = _dataStorageService.GetActiveBit();
            var autoControl = autoList?.FirstOrDefault(c =>
            string.Equals(c.Description, "Auto/Manual", StringComparison.OrdinalIgnoreCase));
            int bitValue = GetPlcDeviceBit("X0"); // Auto/Manual PLC bit

            if (!(autoControl != null && (autoControl.Bit == bitValue)))
            {
                if (autoControl?.Bit == 0 && bitValue == 1)
                    await NotifyOnUIAsync("SOFTWARE IS IN MANUAL MODE. PLEASE SWITCH PLC TO MANUAL MODE.");
                else if (autoControl?.Bit == 1 && bitValue == 0)
                    await NotifyOnUIAsync("SOFTWARE IS IN AUTO MODE. PLEASE SWITCH PLC TO AUTO MODE");
                return;
            }

            var sortedProbeMeasurements = _orderedProbeMeasurements;
            bool firstMeasurementCycle = true;

            do
            {
                // Only check _continueMeasurement before starting a new cycle
                if (!firstMeasurementCycle && !_continueMeasurement)
                {
                    await NotifyOnUIAsync("Stop requested. Finishing current cycle before stopping...");
                    break; // exit loop after finishing current cycle
                }

                // Clear previous readings if not the first cycle
                if (!firstMeasurementCycle)
                {
                    foreach (var probe in sortedProbeMeasurements)
                        probe.Readings.Clear();
                }

                // Wait for robot safe position if not the first cycle
                if (!firstMeasurementCycle)
                {
                    await NotifyOnUIAsync("WAITING FOR ROBOT TO REACH SAFE POSITION...");
                    while (GetPlcDeviceBit("M28") != 0) await Task.Delay(100);
                    await NotifyOnUIAsync("ROBOT IS IN SAFE POSITION. READY TO LOAD NEXT PART.");
                }

                // ===== Start Cycle =====
                if (firstMeasurementCycle)
                {
                    if (autoControl?.Bit == 0)
                    {

                        await NotifyOnUIAsync("LOAD THE PART...");

                        while (GetPlcDeviceBit("X46") != 1) await Task.Delay(100);
                        int Check = GetPlcDeviceBit("M4");
                        if (Check != 1)
                        {
                            MessageBox.Show("SOME CYLINDERS ARE NOT AT HOME!");
                            SetPlcDevice("L25", 1);// hOME BIT FORFCE
                            return;
                        }

                        if (!_continueMeasurement) break;

                        await NotifyOnUIAsync("PRESS THE START SWITCH TO BEGIN MEASUREMENT");
                        while (GetPlcDeviceBit("X1") != 1 && _continueMeasurement) await Task.Delay(100);

                        if (!_continueMeasurement) break;

                        await NotifyOnUIAsync("MOTOR BLOCK CYCLE ON");
                        await WaitForPlcBitAsync("X51");
                        SetPlcDevice("M114", 1);

                        await NotifyOnUIAsync("MOTOR DOWN SIGNAL ON");
                        await WaitForPlcBitAsync("X50");
                        SetPlcDevice("M127", 1);

                        switch (ActiveIdValue)
                        {

                            case 1:
                                await NotifyOnUIAsync("RIGHT_ID_CYL 9.508 AND LEFT_ID_CYL 9.508 ON ");

                                await WaitForPlcBitAsync("X65"); //Right_Block_ID-12.642
                                await WaitForPlcBitAsync("X75"); //Right_Block_ID-8.072
                                await WaitForPlcBitAsync("X71"); //LEFT_BLOCK_ID-12.642
                                await WaitForPlcBitAsync("X101"); //LEFT_BLOCK_ID-8.072
                                SetPlcDevice("M131", 1);//Right_Block_ID-9.508
                                SetPlcDevice("M132", 1);//LEFT_BLOCK_ID-9.508
                                break;
                            case 2:

                                await NotifyOnUIAsync("RIGHT_BLOCK_CYL 12.642 AND LEFT_BLOCK_CYL 12.642 ON");

                                await WaitForPlcBitAsync("X67");
                                await WaitForPlcBitAsync("X61");
                                await WaitForPlcBitAsync("X75");
                                await WaitForPlcBitAsync("X101");
                                await WaitForPlcBitAsync("X63");
                                await WaitForPlcBitAsync("X73");

                                SetPlcDevice("M116", 1);//Right Cylinder ID-12.642
                                SetPlcDevice("M118", 1);//lEFT Cylinder ID-12.642

                                await NotifyOnUIAsync("RIGHT_ID 12.642 AND LEFT_ID 12.642 ON");

                                await WaitForPlcBitAsync("X66");
                                await WaitForPlcBitAsync("X72");

                                SetPlcDevice("M133", 1);//Right Cylinder ID-12.642
                                SetPlcDevice("M134", 1);//lEFT Cylinder ID-12.642

                                break;

                            case 3:
                                await NotifyOnUIAsync("RIGHT_BLOCK_CYL 8.072 AND LEFT_BLOCK_CYL 8.072 ON");

                                await WaitForPlcBitAsync("X77");//Right_CYC_ID-8.072
                                await WaitForPlcBitAsync("X61");//Right_Block_ID-9.508
                                await WaitForPlcBitAsync("X65");//Right_Block_ID-12.642
                                await WaitForPlcBitAsync("X103");//LEFT_CYC_ID-8.072
                                await WaitForPlcBitAsync("X73");//LEFT_BLOCK_ID-12.642
                                await WaitForPlcBitAsync("X63");//LEFT_BLOCK_ID-9.508


                                SetPlcDevice("M120", 1);//Right_Block_ID-8.072 ON
                                SetPlcDevice("M122", 1);//LEFT_BLOCK_ID-8.072

                                await NotifyOnUIAsync("RIGHT_ID 8.072 AND LEFT_ID 8.072 ON");

                                await WaitForPlcBitAsync("X76"); //Right_CYC_ID - 8.072 
                                await WaitForPlcBitAsync("X102");//LEFT_CYC_ID-8.072

                                SetPlcDevice("M135", 1);//Right_ID-8.072 ON
                                SetPlcDevice("M136", 1);//LEFT_ID-8.072 ON

                                break;

                            case 0:
                                // Do nothing
                                break;


                            default:
                                // Optional: handle unexpected values
                                break;
                        }
                        await NotifyOnUIAsync("Starting measurement...");

                    }
                    else
                    {


                        // ================= AUTO MODE PRE-CHECK =================
                        if (GetPlcDeviceBit("M3") != 0)
                        {
                            //MessageBox.Show(
                            //    "Previous cycle is not cleared.\nPlease clear the cycle before starting.",
                            //    "Auto Mode",
                            //    MessageBoxButton.OK,
                            //    MessageBoxImage.Warning);

                            return;
                        }

                        // 1️⃣ Check cylinders at home
                        if (GetPlcDeviceBit("M4") != 1)
                        {
                            MessageBox.Show(
                                "Some cylinders are not at home.\nForcing homing operation.",
                                "Auto Mode",
                                MessageBoxButton.OK,
                                MessageBoxImage.Warning);

                            SetPlcDevice("L25", 1); // Force Home
                            return;
                        }




                        // === AUTO MODE ===
                        if (!_continueMeasurement) break;

                        await NotifyOnUIAsync("PRESS THE ROBO START BUTTON TO BEGIN MEASUREMENT.");
                        while (GetPlcDeviceBit("B2") != 1 && _continueMeasurement) await Task.Delay(100);

                        if (!_continueMeasurement) break;

                        //await NotifyOnUIAsync("Robo Start button pressed");
                        //SetPlcDevice("M301", 1);


                        await NotifyOnUIAsync("WAITING FOR ROBOT TO REACH SAFE POSITION...");

                        //
                        while (GetPlcDeviceBit("M28") != 1) await Task.Delay(100);

                        if (_continueMeasurement)
                            while (GetPlcDeviceBit("X46") != 1) await Task.Delay(100);


                        if (!_continueMeasurement) break;




                        SetPlcDevice("M28", 0);

                        

                        await NotifyOnUIAsync("MOTOR BLOCK CYCLE ON");
                        await WaitForPlcBitAsync("X51");
                        SetPlcDevice("M114", 1);

                        await NotifyOnUIAsync("MOTOR DOWN SIGNAL ON");
                        await WaitForPlcBitAsync("X50");
                        SetPlcDevice("M127", 1);

                        switch (ActiveIdValue)
                        {

                            case 1:
                                await NotifyOnUIAsync("RIGHT_ID_CYL 9.508 AND LEFT_ID_CYL 9.508 ON ");

                                await WaitForPlcBitAsync("X65"); //Right_Block_ID-12.642
                                await WaitForPlcBitAsync("X75"); //Right_Block_ID-8.072
                                await WaitForPlcBitAsync("X71"); //LEFT_BLOCK_ID-12.642
                                await WaitForPlcBitAsync("X101"); //LEFT_BLOCK_ID-8.072
                                SetPlcDevice("M131", 1);//Right_Block_ID-9.508
                                SetPlcDevice("M132", 1);//LEFT_BLOCK_ID-9.508
                                break;
                            case 2:

                                await NotifyOnUIAsync("RIGHT_BLOCK_CYL 12.642 AND LEFT_BLOCK_CYL 12.642 ON");

                                await WaitForPlcBitAsync("X67");
                                await WaitForPlcBitAsync("X61");
                                await WaitForPlcBitAsync("X75");
                                await WaitForPlcBitAsync("X101");
                                await WaitForPlcBitAsync("X63");
                                await WaitForPlcBitAsync("X73");

                                SetPlcDevice("M116", 1);//Right Cylinder ID-12.642
                                SetPlcDevice("M118", 1);//lEFT Cylinder ID-12.642

                                await NotifyOnUIAsync("RIGHT_ID 12.642 AND LEFT_ID 12.642 ON");

                                await WaitForPlcBitAsync("X66");
                                await WaitForPlcBitAsync("X72");

                                SetPlcDevice("M133", 1);//Right Cylinder ID-12.642
                                SetPlcDevice("M134", 1);//lEFT Cylinder ID-12.642

                                break;

                            case 3:
                                await NotifyOnUIAsync("RIGHT_BLOCK_CYL 8.072 AND LEFT_BLOCK_CYL 8.072 ON");

                                await WaitForPlcBitAsync("X77");//Right_CYC_ID-8.072
                                await WaitForPlcBitAsync("X61");//Right_Block_ID-9.508
                                await WaitForPlcBitAsync("X65");//Right_Block_ID-12.642
                                await WaitForPlcBitAsync("X103");//LEFT_CYC_ID-8.072
                                await WaitForPlcBitAsync("X73");//LEFT_BLOCK_ID-12.642
                                await WaitForPlcBitAsync("X63");//LEFT_BLOCK_ID-9.508


                                SetPlcDevice("M120", 1);//Right_Block_ID-8.072 ON
                                SetPlcDevice("M122", 1);//LEFT_BLOCK_ID-8.072

                                await NotifyOnUIAsync("RIGHT_ID 8.072 AND LEFT_ID 8.072 ON");

                                await WaitForPlcBitAsync("X76"); //Right_CYC_ID - 8.072 
                                await WaitForPlcBitAsync("X102");//LEFT_CYC_ID-8.072

                                SetPlcDevice("M135", 1);//Right_ID-8.072 ON
                                SetPlcDevice("M136", 1);//LEFT_ID-8.072 ON

                                break;

                            case 0:
                                // Do nothing
                                break;


                            default:
                                // Optional: handle unexpected values
                                break;
                        }
                    }
                }
                else
                {
                    if (autoControl?.Bit == 0)
                    {

                        //await NotifyOnUIAsync(" Remove the Part");
                        if (!_continueMeasurement) break;

                        // Wait until part is removed (X46 = 0)
                        while (GetPlcDeviceBit("X46") == 1)
                        {
                            await NotifyOnUIAsync("REMOVE THE PART");
                            await Task.Delay(300); // avoid CPU overload
                        }

                        //  Part removed → continue further logic here



                        if (!_continueMeasurement)
                            break;

                        // 1️⃣ Ask operator to load the part
                        await NotifyOnUIAsync("LOAD THE PART..");
                        while (GetPlcDeviceBit("X46") != 1) await Task.Delay(100);
                        int Check = GetPlcDeviceBit("M4");
                        if (Check != 1)
                        {
                            MessageBox.Show("SOME CYLINDERS ARE NOT AT HOME!");
                            SetPlcDevice("L25", 1);// hOME BIT FORFCE
                            return;
                        }

                        if (!_continueMeasurement)
                            break;

                        // 2️⃣ Ask operator to press start switch
                        await NotifyOnUIAsync("PART DETECTED. PRESS START SWITCH TO BEGIN MEASUREMENT.");
                        while (GetPlcDeviceBit("X1") != 1 && _continueMeasurement)
                            await Task.Delay(100);

                        ResetAll();
                        if (!_continueMeasurement)
                            break;

                        await NotifyOnUIAsync("MOTOR BLOCK CYCLE ACTIVATED");
                        await WaitForPlcBitAsync("X51");
                        SetPlcDevice("M114", 1);

                        await NotifyOnUIAsync("MOTOR DOWN SIGNAL ACTIVE");
                        await WaitForPlcBitAsync("X50");
                        SetPlcDevice("M127", 1);

                        switch (ActiveIdValue)
                        {

                            case 1:
                                await NotifyOnUIAsync("RIGHT_BLOCK_ID = 9.508 AND LEFT_BLOCK_ID = 9.508");

                                await WaitForPlcBitAsync("X65"); //Right_Block_ID-12.642
                                await WaitForPlcBitAsync("X75"); //Right_Block_ID-8.072
                                await WaitForPlcBitAsync("X71"); //LEFT_BLOCK_ID-12.642
                                await WaitForPlcBitAsync("X101"); //LEFT_BLOCK_ID-8.072
                                SetPlcDevice("M131", 1);//Right_Block_ID-9.508
                                SetPlcDevice("M132", 1);//LEFT_BLOCK_ID-9.508
                                break;
                            case 2:

                                await NotifyOnUIAsync("RIGHT_BLOCK_CYL 12.642 AND LEFT_BLOCK_CYL 12.642 ON");

                                await WaitForPlcBitAsync("X67");
                                await WaitForPlcBitAsync("X61");
                                await WaitForPlcBitAsync("X75");
                                await WaitForPlcBitAsync("X101");
                                await WaitForPlcBitAsync("X63");
                                await WaitForPlcBitAsync("X73");

                                SetPlcDevice("M116", 1);//Right Cylinder ID-12.642
                                SetPlcDevice("M118", 1);//lEFT Cylinder ID-12.642

                                await NotifyOnUIAsync("RIGHT_ID 12.642 AND LEFT_ID 12.642 ON");

                                await WaitForPlcBitAsync("X66");
                                await WaitForPlcBitAsync("X72");

                                SetPlcDevice("M133", 1);//Right Cylinder ID-12.642
                                SetPlcDevice("M134", 1);//lEFT Cylinder ID-12.642

                                break;

                            case 3:
                                await NotifyOnUIAsync("RIGHT_BLOCK_CYL 8.072 AND LEFT_BLOCK_CYL 8.072 ON");

                                await WaitForPlcBitAsync("X77");//Right_CYC_ID-8.072
                                await WaitForPlcBitAsync("X61");//Right_Block_ID-9.508
                                await WaitForPlcBitAsync("X65");//Right_Block_ID-12.642
                                await WaitForPlcBitAsync("X103");//LEFT_CYC_ID-8.072
                                await WaitForPlcBitAsync("X73");//LEFT_BLOCK_ID-12.642
                                await WaitForPlcBitAsync("X63");//LEFT_BLOCK_ID-9.508


                                SetPlcDevice("M120", 1);//Right_Block_ID-8.072 ON
                                SetPlcDevice("M122", 1);//LEFT_BLOCK_ID-8.072

                                await NotifyOnUIAsync("RIGHT_ID 8.072 AND LEFT_ID 8.072 ON");

                                await WaitForPlcBitAsync("X76"); //Right_CYC_ID - 8.072 
                                await WaitForPlcBitAsync("X102");//LEFT_CYC_ID-8.072

                                SetPlcDevice("M135", 1);//Right_ID-8.072 ON
                                SetPlcDevice("M136", 1);//LEFT_ID-8.072 ON

                                break;

                            case 0:
                                // Do nothing
                                break;


                            default:
                                // Optional: handle unexpected values
                                break;
                        }
                        await NotifyOnUIAsync("Starting measurement..");

                    }

                    else
                    {

                        // === AUTO MODE REPEAT ===
                        if (!_continueMeasurement) break;


                        await NotifyOnUIAsync("ROBOT IS IN SAFE POSITION. READY TO LOAD NEXT PART.");
                        while (GetPlcDeviceBit("M28") != 1) await Task.Delay(100);


                        if (_continueMeasurement)
                            while (GetPlcDeviceBit("X46") != 1) await Task.Delay(1000);


                        if (!_continueMeasurement) break;

                        ResetAll();

                        SetPlcDevice("M28", 0);

                        if (!_continueMeasurement) break;


                        await NotifyOnUIAsync("MOTOR BLOCK CYCLE ON");
                        await WaitForPlcBitAsync("X51");
                        SetPlcDevice("M114", 1);

                        await NotifyOnUIAsync("MOTOR DOWN SIGNAL ON");
                        await WaitForPlcBitAsync("X50");
                        SetPlcDevice("M127", 1);

                        switch (ActiveIdValue)
                        {

                            case 1:
                                await NotifyOnUIAsync("RIGHT_BLOCK_ID = 9.508 AND LEFT_BLOCK_ID = 9.508");

                                await WaitForPlcBitAsync("X65"); //Right_Block_ID-12.642
                                await WaitForPlcBitAsync("X75"); //Right_Block_ID-8.072
                                await WaitForPlcBitAsync("X71"); //LEFT_BLOCK_ID-12.642
                                await WaitForPlcBitAsync("X101"); //LEFT_BLOCK_ID-8.072
                                SetPlcDevice("M131", 1);//Right_Block_ID-9.508
                                SetPlcDevice("M132", 1);//LEFT_BLOCK_ID-9.508
                                break;
                            case 2:

                                await NotifyOnUIAsync("RIGHT_BLOCK_CYL 12.642 AND LEFT_BLOCK_CYL 12.642 ON");

                                await WaitForPlcBitAsync("X67");
                                await WaitForPlcBitAsync("X61");
                                await WaitForPlcBitAsync("X75");
                                await WaitForPlcBitAsync("X101");
                                await WaitForPlcBitAsync("X63");
                                await WaitForPlcBitAsync("X73");

                                SetPlcDevice("M116", 1);//Right Cylinder ID-12.642
                                SetPlcDevice("M118", 1);//lEFT Cylinder ID-12.642

                                await NotifyOnUIAsync("RIGHT_ID 12.642 AND LEFT_ID 12.642 ON");

                                await WaitForPlcBitAsync("X66");
                                await WaitForPlcBitAsync("X72");

                                SetPlcDevice("M133", 1);//Right Cylinder ID-12.642
                                SetPlcDevice("M134", 1);//lEFT Cylinder ID-12.642

                                break;

                            case 3:
                                await NotifyOnUIAsync("RIGHT_BLOCK_CYL 8.072 AND LEFT_BLOCK_CYL 8.072 ON");

                                await WaitForPlcBitAsync("X77");//Right_CYC_ID-8.072
                                await WaitForPlcBitAsync("X61");//Right_Block_ID-9.508
                                await WaitForPlcBitAsync("X65");//Right_Block_ID-12.642
                                await WaitForPlcBitAsync("X103");//LEFT_CYC_ID-8.072
                                await WaitForPlcBitAsync("X73");//LEFT_BLOCK_ID-12.642
                                await WaitForPlcBitAsync("X63");//LEFT_BLOCK_ID-9.508


                                SetPlcDevice("M120", 1);//Right_Block_ID-8.072 ON
                                SetPlcDevice("M122", 1);//LEFT_BLOCK_ID-8.072

                                await NotifyOnUIAsync("RIGHT_ID 8.072 AND LEFT_ID = 8.072 ON");

                                await WaitForPlcBitAsync("X76"); //Right_CYC_ID - 8.072 
                                await WaitForPlcBitAsync("X102");//LEFT_CYC_ID-8.072

                                SetPlcDevice("M135", 1);//Right_ID-8.072 ON
                                SetPlcDevice("M136", 1);//LEFT_ID-8.072 ON

                                break;

                            case 0:
                                // Do nothing
                                break;


                            default:
                                // Optional: handle unexpected values
                                break;
                        }
                    }
                }

                if (_continueMeasurement)
                {
                    await RunMotorAndCollectReadingsAsync(sortedProbeMeasurements, ProcedureMode.Measurement);
                }



                var probeMeasurementByNameMeasurement = sortedProbeMeasurements
                    .Where(pm => !string.IsNullOrEmpty(pm.Name))
                    .ToDictionary(pm => pm.Name!);

                HandleMeasurementStage(probeMeasurementByNameMeasurement, _currentPartCode);


                firstMeasurementCycle = false;

            } while (ShouldContinueMeasurement());

            // ===== Stop & Reset =====
            await NotifyOnUIAsync("Measurement cycle stopped by user.");

            //SetPlcDevice("M10", 0);   // Motor OFF
            //SetPlcDevice("M300", 0);  // Measurement Complete OFF
            MeasurementStopped?.Invoke();
        }



        private async Task RunMotorAndCollectReadingsAsync(List<ProbeMeasurement> sortedProbeMeasurements,ProcedureMode mode)
        {
            const int StabilizationDelayMs = 700;
            int SamplesPerProbe =ArraySize;
            const int PollIntervalMs = 50;
            const int InitialDiscardSamples = 3; // kept, not used
            const int TrimCount = 3;

            var probesByName = sortedProbeMeasurements.ToDictionary(p => p.Name);

            // OD probes combined with RN readings
            var odGroupsByName = new Dictionary<string, List<string>>
            {
                { "OD1", new List<string> { "RN1" } },
                { "OD2", new List<string> { "RN2" } },
                { "OD3", new List<string> { "RN3" } },
                { "OD4", new List<string> { "RN4" } },
                { "OD5", new List<string> { "RN5" } },
                { "ID-1", new List<string> { "RN6" } },
                { "ID-2", new List<string> { "RN7" } }
            };



            // 1️⃣ Check cylinders at home
            if (GetPlcDeviceBit("X6") != 0)
            {
                MessageBox.Show(
                    "Probe Pressure is Low...\nCheck the Air Pressure!!",
                    "Auto Mode",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                
                //SetPlcDevice("L25", 1); // Force Home
                return;
            }

            try
            {
                await _plcProbeService.LoadProbesAsync(_currentPartCode);

                if (!await _plcProbeService.ConnectSerialAsync())
                    return;


                await NotifyOnUIAsync("PROBES ON");
                await Task.Delay(StabilizationDelayMs);
                SetPlcDevice("M137", 1);

                await NotifyOnUIAsync("SYSTEM IS COLLECTING READINGS.");
                await Task.Delay(500);
                SetPlcDevice("M101", 1);

                // Reset all probes
                foreach (var pm in sortedProbeMeasurements)
                {
                    pm.Readings.Clear();
                    pm.MinValue = double.MaxValue;
                    pm.MaxValue = double.MinValue;
                }   
                await Task.Delay(10);
                _plcProbeService.StartLiveReading(8, PollIntervalMs);

                var probesPending = sortedProbeMeasurements
                    .ToDictionary(p => p.Name, _ => 0);

                // ============================
                // STEP 1: COLLECT ALL READINGS (NO DISCARD)
                // ============================
                while (!Abort && probesPending.Values.Any(v => v < SamplesPerProbe))
                {
                    foreach (var pm in sortedProbeMeasurements)
                    {
                        if (probesPending[pm.Name] >= SamplesPerProbe)
                            continue;

                        if (_plcProbeService.ProbeReadings.TryGetValue(pm.ProbeId, out var reading))
                        {
                            double value = reading.Value;

                            if (double.IsNaN(value) || value == 0)
                                continue;

                            probesPending[pm.Name]++;
                            pm.Readings.Add(Math.Round(value, 3)); // ✅ store all
                        }
                    }

                    await Task.Delay(PollIntervalMs);
                }
                await NotifyOnUIAsync("READINGS DONE.");

                _plcProbeService.StopAndCloseSerial();
                SetPlcDevice("M101", 0);
                SetPlcDevice("M137", 0);
                SetPlcDevice("L25", 1);

                // =====================================================
                // STEP 2: ADD RN VALUES TO OD PROBES (ALIGN EXACTLY)
                // =====================================================


                foreach (var pm in sortedProbeMeasurements)
                {

                    for (int i = 0; i < pm.Readings.Count; i++)
                    {
                      var pr=  pm.Readings[i];
                    }

                }

                foreach (var od in odGroupsByName)
                {
                    if (!probesByName.TryGetValue(od.Key, out var odProbe))
                        continue;

                    var sourceProbes = od.Value
                        .Where(name => probesByName.ContainsKey(name))
                        .Select(name => probesByName[name])
                        .ToList();

                    if (sourceProbes.Count == 0)
                        continue;

                    int sampleCount = sourceProbes.Min(p => p.Readings.Count);

                    while (odProbe.Readings.Count < sampleCount)
                        odProbe.Readings.Add(0);

                    for (int i = 0; i < sampleCount; i++)
                    {
                        double rnSum = 0;
                        foreach (var src in sourceProbes)
                            rnSum += src.Readings[i];

                        odProbe.Readings[i] = Math.Round(odProbe.Readings[i] + rnSum, 3);
                    }

                    if (odProbe.Readings.Count > sampleCount)
                        odProbe.Readings = odProbe.Readings.Take(sampleCount).ToList();
                }

                // ============================
                // STEP 3: TRIM FIRST & LAST (TIME-BASED)
                // ============================
                foreach (var pm in sortedProbeMeasurements)
                {
                    if (pm.Readings.Count <= TrimCount * 2)
                        continue;

                    pm.Readings = pm.Readings
                        .Skip(TrimCount)
                        .Take(pm.Readings.Count - (TrimCount * 2))
                        .ToList();
                }

                // ============================
                // STEP 4: SORT & MIN / MAX
                // ============================
                foreach (var pm in sortedProbeMeasurements)
                {
                    if (pm.Readings.Count == 0)
                        continue;

                    pm.Readings = pm.Readings.OrderBy(v => v).ToList();
                    pm.MinValue = pm.Readings.First();
                    pm.MaxValue = pm.Readings.Last();
                }

                //Check total Sorted Readings for all probes /Single also
                foreach (var pm in sortedProbeMeasurements)
                {

                    for (int i = 0; i < pm.Readings.Count; i++)
                    {
                        var pr = pm.Readings[i];
                    }

                }
            }
            catch
            {
                _plcProbeService.StopAndCloseSerial();
                SetPlcDevice("M101", 0);
                SetPlcDevice("M137", 0);
                SetPlcDevice("L25", 1);
                throw;
            }
        }



        private async Task HandleProcedurePostProcessingAsync(ProcedureMode mode, List<ProbeMeasurement> probeMeasurements, bool isFirstMeasurementCycle)
        {
            switch (mode)
            {
                case ProcedureMode.Mastering:
                    await NotifyOnUIAsync("MASTERING COMPLETED. PRESS ENTER TO INSPECT THE MASTER.");

                    var masterValues = probeMeasurements
                             .Where(pm => !string.IsNullOrEmpty(pm.ProbeId))
                             .ToDictionary(
                                 pm => pm.ProbeId!,
                                 pm => pm.MaxValue);   // only Max

                    OnCalculatedValuesReady(masterValues);   // OK: Dictionary<string, double>



                    OnCalculatedValuesReady(masterValues);
                    await HandleMasteringStageAsync();
                    break;

                case ProcedureMode.MasterInspection:
                    {
                        await NotifyOnUIAsync("Master Inspection Completed...");

                        // Build dictionary safely (avoid duplicate keys)
                        var probeMeasurementByName = probeMeasurements
                            .Where(pm => !string.IsNullOrEmpty(pm.Name))
                            .GroupBy(pm => pm.Name)
                            .ToDictionary(g => g.Key, g => g.First());

                        // Perform master inspection
                        HandleMasterInspectionStage(probeMeasurementByName, _currentPartCode);

                        // 🔹 Read Auto/Manual PLC Bit (X14)
                        int bitValue = GetPlcDeviceBit("X0");  // 1 = Auto mode

                        if (bitValue == 1)
                        {
                            // 🔸 Signal for robo safe position
                            Thread.Sleep(1500);

                            SetPlcDevice("M19", 1); // trigger robot to move to safe position
                            //await NotifyOnUIAsync("Waiting For Robot to Unload the Master");

                            while (GetPlcDeviceBit("M16") != 0) await Task.Delay(100);

                        }
                        else
                        {
                            // 🔸 Skip PLC trigger if in Manual Mode
                            await NotifyOnUIAsync("Master Inspection Completed (Manual Mode)");
                        }

                        break;
                    }


            }
        }






        private async Task HandleMasteringStageAsync()
        {
            if (Abort)
            {
                MessageBox.Show("Mastering aborted.", "Abort", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var effectiveByName = BuildEffectiveProbesByName(_orderedProbeMeasurements)
                                .ToDictionary(k => k.Key, v => v.Value);

            // key = ProbeId, value = (Min, Max)
            var masterValues = effectiveByName.ToDictionary(
                kvp => kvp.Value.ProbeId,
                kvp => (Min: kvp.Value.MinValue,
                        Max: kvp.Value.MaxValue));

            _dataStorageService.SaveProbeReadings(
                _dataStorageService.GetProbeInstallByPartNumber(_currentPartCode),
                _currentPartCode,
                masterValues);

            MasterComplete = true;
            var resultsWithStatus = effectiveByName.Values
                .ToDictionary(
                    pm => pm.Name!,  // ✅ Display name for UI
                    pm => new ParameterResult { Value = pm.MaxValue, IsOk = true });

            OnCalculatedValuesWithStatusReady(resultsWithStatus);

            // Determine current mode
            var autoList = _dataStorageService.GetActiveBit();
            var autoControl = autoList?.FirstOrDefault(c => string.Equals(c.Description, "Auto/Manual", StringComparison.OrdinalIgnoreCase));
            int bitValue = GetPlcDeviceBit("X0"); // PLC Auto/Manual bit

            bool isAuto = autoControl != null && autoControl.Bit == 1 && bitValue == 1;

            if (isAuto)
            {
                Thread.Sleep(1500);

                SetPlcDevice("M19", 1); // trigger robot to move to safe position
                await NotifyOnUIAsync("Waiting Safe position from Robot...");

                while (GetPlcDeviceBit("M16") != 0) await Task.Delay(100);

            }

            //SetPlcDevice("L25", 1);
            // Mastering complete message
            await NotifyOnUIAsync("Mastering completed. Press Enter to inspect the master.");
        }
        

        private bool IsValidParameter(string param, Dictionary<string, ProbeMeasurement> probeMeasurements, Dictionary<string, double> dbRefDict)
        {
            return probeMeasurements.ContainsKey(param) && dbRefDict.ContainsKey(param);
        }
        private void HandleMasterInspectionStage(
            Dictionary<string, ProbeMeasurement> probeMeasurements,
            string partCode)
        {
            var probeMeasurementByName =
                BuildEffectiveProbesByName(probeMeasurements.Values.ToList());

            var dbRefList = _dataStorageService.GetMasterProbeRef(_currentPartCode);
            var mode = ProcedureMode.MasterInspection;

            var dbRefDict = dbRefList
                .GroupBy(x => x.Name)
                .ToDictionary(
                    g => g.Key,
                    g => (Min: g.First().Min, Max: g.First().Max));

            var masterVals = _dataStorageService.GetMasterReadingByPart(partCode);

            var parameterNames = masterVals
                .Select(m => m.Parameter)
                .Distinct()
                .ToList();

            // 🔥 Now stores Min/Max
            var calculatedValues = new Dictionary<string, (double Min, double Max)>();

            var parameterInfos = parameterNames.Select(p => new ParameterInfo
            {
                Name = p,
                SignChange = 0,
                Compensation = 0
            }).ToList();

            foreach (var pInfo in parameterInfos)
            {
                try
                {
                    calculatedValues[pInfo.Name] =
                        CalculateProbeValue(pInfo, probeMeasurementByName, dbRefDict, mode);
                }
                catch
                {
                    calculatedValues[pInfo.Name] = (double.NaN, double.NaN);
                }
            }

            bool overallOk = true;
            var resultsWithStatus = new Dictionary<string, ParameterResult>();

            foreach (var paramName in parameterNames)
            {
                var measured = calculatedValues[paramName];

                var config = masterVals.FirstOrDefault(m =>
                    string.Equals(m.Para_No, paramName, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(m.Parameter, paramName, StringComparison.OrdinalIgnoreCase));

                double masterVal = config?.Nominal ?? 0;
                double tolPlus = config?.RTolPlus ?? 0;
                double tolMinus = config?.RTolMinus ?? 0;

                double minAllowed = masterVal - tolMinus;
                double maxAllowed = masterVal + tolPlus;

                double measuredMin = measured.Min;
                double measuredMax = measured.Max;

                bool isOk =
                    !double.IsNaN(measuredMax) && measuredMax >= minAllowed && measuredMax <= maxAllowed;
                if (!isOk)
                    overallOk = false;

                resultsWithStatus[paramName] = new ParameterResult
                {
                    Min = Math.Round(measuredMin, 3),
                    Value = Math.Round(measuredMax, 3),
                    IsOk = isOk
                };
            }

            MasteringOK = overallOk;
            OnCalculatedValuesWithStatusReady(resultsWithStatus);
        }



        private void HandleMeasurementStage(Dictionary<string, ProbeMeasurement> probeMeasurements, string partCode)
        {
            // Get reference master values
            var dbRefList = _dataStorageService.GetMasterProbeRef(partCode);
            var mode = ProcedureMode.Measurement;
            var dbRefDict = dbRefList
                 .GroupBy(x => x.Name)
                 .ToDictionary(
                     g => g.Key,
                     g => (Min: g.First().Min, Max: g.First().Max));

            // Load master part configurations (nominal values and tolerances)
            var masterVals = _dataStorageService.GetPartConfig(partCode);

            // 17 measurement parameters in order

            var parameterNames = masterVals
                .Select(m => m.Parameter)
                .Distinct()
                .ToList();
            //string[] parameterNames = new string[]
            //{
            //        "Overall Length", "Datum to End", "Head Diameter", "Groove Position",
            //        "Stem Dia Near Groove", "Stem Dia Near Undercut", "Groove Diameter",
            //        "Straightness", "Seat Height", "Seat Runout", "Datum to Groove",
            //        "Ovality SDG", "Ovality SDU", "Ovality Head", "Stem Taper",
            //        "Face Runout", "End Face Runout"
            //};

            var parameterInfos = masterVals
        .GroupBy(m => m.Parameter)
        .Select(g => new ParameterInfo
        {
            Name = g.Key,
            SignChange = g.First().Sign_Change,
            Compensation = g.First().Compensation
        })
        .ToList();

            var calculatedValues = new Dictionary<string, (double Min, double Max)>();

            // Calculate each probe value
            foreach (var pInfo in parameterInfos)
            {
                try
                {
                    calculatedValues[pInfo.Name] =
                        CalculateProbeValue(pInfo, probeMeasurements, dbRefDict, mode);
                }
                catch
                {
                    calculatedValues[pInfo.Name] = (double.NaN, double.NaN);
                }
            }


            // Determine overall OK/NG and prepare results
            bool overallOk = true;
            var resultsWithStatus = new Dictionary<string, ParameterResult>();

            foreach (var paramName in parameterNames)
            {
                var measured = calculatedValues[paramName];

                var config = masterVals?.FirstOrDefault(m =>
                    string.Equals(m.Para_No, paramName, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(m.Parameter, paramName, StringComparison.OrdinalIgnoreCase)
                );

                double masterVal = config?.Nominal ?? 0;
                double tolPlus = config?.RTolPlus ?? 0;
                double tolMinus = config?.RTolMinus ?? 0;

                double minAllowed = masterVal - tolMinus;
                double maxAllowed = masterVal + tolPlus;

                double measuredMin = measured.Min;
                double measuredMax = measured.Max;

                bool isOk =
                    !double.IsNaN(measuredMax) && measuredMax >= minAllowed && measuredMax <= maxAllowed;

                if (!isOk)
                    overallOk = false;

                resultsWithStatus[paramName] = new ParameterResult
                {
                    Min = Math.Round(measuredMin, 3),   // ✅ MEASURED MIN
                    Value = Math.Round(measuredMax, 3),   // ✅ MEASURED MAX
                    IsOk = isOk
                };
            }


            MasteringOK = overallOk;

            // Send calculated results to the Result Page
            OnCalculatedValuesWithStatusReady(resultsWithStatus);




            int bitValue = GetPlcDeviceBit("X0"); // Auto/Manual PLC bit

            //var autoList = _dataStorageService.GetActiveBit();

            //var shControl = autoList.FirstOrDefault(c => c.Code == "SH");
            //var sroControl = autoList.FirstOrDefault(c => c.Code == "SRO");
            //var stdiControl = autoList.FirstOrDefault(c => c.Code == "STDI");
            //var gdControl = autoList.FirstOrDefault(c => c.Code == "GD");

            try
            {
                if (bitValue == 1) // Auto mode only
                {
                    int ngCount = resultsWithStatus.Count(r => !r.Value.IsOk);

                    // If any NG exists, pulse signal M13
                    if (ngCount > 0)
                    {
                        SetPlcDevice("M13", 1); // Turn ON
                        Thread.Sleep(10);        // 5 ms pulse (adjust as needed)
                        //SetPlcDevice("M13", 0);
                       // SetPlcDevice("M30", 1); // OK

                        SetPlcDevice("M31",1);
                        // Turn OFF
                    }


                    // ---- Case 1: More than 2 NGs => Direct general rejection ----
                    if (ngCount==0 )
                    {
                        SetPlcDevice("M30", 1); // OK
                        return; // stop checking further
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error setting PLC bits for rejection: {ex.Message}");
            }
            }



        

        protected virtual void OnCalculatedValuesReady(Dictionary<string, double> calculatedValues)
        {
            CalculatedValuesReady?.Invoke(this, calculatedValues);
        }



        // Calculation dispatcher adapted to accept both probeMeasurements and dbRefDict
        private (double Min, double Max) CalculateProbeValue(
     ParameterInfo paramInfo,
     Dictionary<string, ProbeMeasurement> probeMeasurements,
     Dictionary<string, (double Min, double Max)> dbRefDict,
     ProcedureMode mode)
        {
            string paramName = paramInfo.Name.ToLower();
            int signChange = paramInfo.SignChange;
            double compensation = paramInfo.Compensation;

            switch (paramName)
            {
                case "od1":
                    return CalculateOD1(probeMeasurements, dbRefDict, mode, signChange, compensation);

                case "od2":
                    return CalculateOD2(probeMeasurements, dbRefDict, mode, signChange, compensation);

                case "od3":
                    return CalculateOD3(probeMeasurements, dbRefDict, mode, signChange, compensation);

                case "od4":
                    return CalculateOD4(probeMeasurements, dbRefDict, mode, signChange, compensation);

                case "od5":
                    return CalculateOD5(probeMeasurements, dbRefDict, mode, signChange, compensation);

                case "id-1":
                    return CalculateID1(probeMeasurements, dbRefDict, mode, signChange, compensation, ActiveIdValue);

                case "id-2":
                    return CalculateID2(probeMeasurements, dbRefDict, mode, signChange, compensation, ActiveIdValue);

                case "ol":
                    {
                        var val = CalculateOverallLength(probeMeasurements, dbRefDict, mode, signChange, compensation);
                        return (val, val); // OL is a single value → promote to range
                    }

                case "rn1":
                    {
                        var val = CalculateStepRunout1(probeMeasurements, dbRefDict, mode, signChange, compensation);
                        return (val, val);
                    }

                case "rn2":
                    {
                        var val = CalculateStepRunout2(probeMeasurements, dbRefDict, mode, signChange, compensation);
                        return (val, val);
                    }

                case "rn3":
                    {
                        var val = CalculateRN1(probeMeasurements, dbRefDict, mode, signChange, compensation);
                        return (val, val);
                    }

                case "rn4":
                    {
                        var val = CalculateRN2(probeMeasurements, dbRefDict, mode, signChange, compensation);
                        return (val, val);
                    }

                case "rn5":
                    {
                        var val = CalculateRN3(probeMeasurements, dbRefDict, mode, signChange, compensation);
                        return (val, val);
                    }

                case "rn6":
                    {
                        var val = CalculateRN4(probeMeasurements, dbRefDict, mode, signChange, compensation);
                        return (val, val);
                    }

                case "rn7":
                    {
                        var val = CalculateRN5(probeMeasurements, dbRefDict, mode, signChange, compensation);
                        return (val, val);
                    }

                default:
                    Debug.WriteLine($"⚠️ Unknown parameter: {paramInfo.Name}");
                    return (double.NaN, double.NaN);
            }

        }

        #region 🔥 MAIN PARAMETERS (Summing Live & Master from 2 Probes)


        private (double Min, double Max) CalculateOD1(
     Dictionary<string, ProbeMeasurement> probeMeasurements,
     Dictionary<string, (double Min, double Max)> dbRefDict,
     ProcedureMode mode,
     int signChange,
     double compensation)
        {
            if (!probeMeasurements.TryGetValue("OD1", out var pm))
                return (double.NaN, double.NaN);

            double liveValueMax = pm.Readings.Any() ? pm.MaxValue : 0;
            double liveValueMin = pm.Readings.Any() ? pm.MinValue : 0;

            if (liveValueMax == 0)
                return (double.NaN, double.NaN);

            if (!dbRefDict.TryGetValue("OD1", out var refVal))
                return (double.NaN, double.NaN);

            // ---------- MIN calculation ----------
            double offsetMin = liveValueMin - refVal.Min;
            double minResult;

            if (mode == ProcedureMode.Measurement)
            {
                minResult = (signChange == 1)
                    ? pm.MasterValue - offsetMin
                    : pm.MasterValue + offsetMin;

                if (compensation != 0)
                    minResult += compensation;
            }
            else
            {
                minResult = pm.MasterValue + offsetMin;
            }

            // ---------- MAX calculation ----------
            double offsetMax = liveValueMax - refVal.Max;
            double maxResult;

            if (mode == ProcedureMode.Measurement)
            {
                maxResult = (signChange == 1)
                    ? pm.MasterValue - offsetMax
                    : pm.MasterValue + offsetMax;

                if (compensation != 0)
                    maxResult += compensation;
            }
            else
            {
                maxResult = pm.MasterValue + offsetMax;
            }

            return (
                Math.Round(minResult, 3),
                Math.Round(maxResult, 3)
            );
        }



        private (double Min, double Max) CalculateOD2(
    Dictionary<string, ProbeMeasurement> probeMeasurements,
    Dictionary<string, (double Min, double Max)> dbRefDict,
    ProcedureMode mode,
    int signChange,
    double compensation)
        {
            if (!probeMeasurements.TryGetValue("OD2", out var pm))
                return (double.NaN, double.NaN);

            double liveValueMax = pm.Readings.Any() ? pm.MaxValue : 0;
            double liveValueMin = pm.Readings.Any() ? pm.MinValue : 0;
            if (liveValueMax == 0)
                return (double.NaN, double.NaN);

            if (!dbRefDict.TryGetValue("OD2", out var db))
                return (double.NaN, double.NaN);

            double offsetMin = liveValueMin  - db.Min;
            double offsetMax = liveValueMax - db.Max;

            double min = signChange == 1 ? pm.MasterValue - offsetMin : pm.MasterValue + offsetMin;
            double max = signChange == 1 ? pm.MasterValue - offsetMax : pm.MasterValue + offsetMax;

            if (mode == ProcedureMode.Measurement && compensation != 0)
            {
                min += compensation;
                max += compensation;
            }

            return (Math.Round(min, 3), Math.Round(max, 3));
        }

        private (double Min, double Max) CalculateOD3(
     Dictionary<string, ProbeMeasurement> probeMeasurements,
     Dictionary<string, (double Min, double Max)> dbRefDict,
     ProcedureMode mode,
     int signChange,
     double compensation)
        {
            if (!probeMeasurements.TryGetValue("OD3", out var pm))
                return (double.NaN, double.NaN);

            double liveValueMax = pm.Readings.Any() ? pm.MaxValue : 0;
            double liveValueMin = pm.Readings.Any() ? pm.MinValue : 0;
            if (liveValueMax == 0)
                return (double.NaN, double.NaN);

            if (!dbRefDict.TryGetValue("OD3", out var db))
                return (double.NaN, double.NaN);

            double offsetMin = liveValueMin - db.Min;
            double offsetMax = liveValueMax - db.Max;


            double min = signChange == 1 ? pm.MasterValue - offsetMin : pm.MasterValue + offsetMin;
            double max = signChange == 1 ? pm.MasterValue - offsetMax : pm.MasterValue + offsetMax;

            if (mode == ProcedureMode.Measurement && compensation != 0)
            {
                min += compensation;
                max += compensation;
            }

            return (Math.Round(min, 3), Math.Round(max, 3));
        }


        private (double Min, double Max) CalculateOD4(
     Dictionary<string, ProbeMeasurement> probeMeasurements,
     Dictionary<string, (double Min, double Max)> dbRefDict,
     ProcedureMode mode,
     int signChange,
     double compensation)
        {
            if (!probeMeasurements.TryGetValue("OD4", out var pm))
                return (double.NaN, double.NaN);

            double liveValueMax = pm.Readings.Any() ? pm.MaxValue : 0;
            double liveValueMin = pm.Readings.Any() ? pm.MinValue : 0;
            if (liveValueMax == 0)
                return (double.NaN, double.NaN);

            if (!dbRefDict.TryGetValue("OD4", out var db))
                return (double.NaN, double.NaN);

            double offsetMin = liveValueMin - db.Min;
            double offsetMax = liveValueMax - db.Max;

            double min = signChange == 1 ? pm.MasterValue - offsetMin : pm.MasterValue + offsetMin;
            double max = signChange == 1 ? pm.MasterValue - offsetMax : pm.MasterValue + offsetMax;

            if (mode == ProcedureMode.Measurement && compensation != 0)
            {
                min += compensation;
                max += compensation;
            }

            return (Math.Round(min, 3), Math.Round(max, 3));
        }


        private (double Min, double Max) CalculateOD5(
     Dictionary<string, ProbeMeasurement> probeMeasurements,
     Dictionary<string, (double Min, double Max)> dbRefDict,
     ProcedureMode mode,
     int signChange,
     double compensation)
        {
            if (!probeMeasurements.TryGetValue("OD5", out var pm))
                return (double.NaN, double.NaN);

            double liveValueMax = pm.Readings.Any() ? pm.MaxValue : 0;
            double liveValueMin = pm.Readings.Any() ? pm.MinValue : 0;
            if (liveValueMax == 0)
                return (double.NaN, double.NaN);

            if (!dbRefDict.TryGetValue("OD5", out var db))
                return (double.NaN, double.NaN);

            double offsetMin = liveValueMin - db.Min;
            double offsetMax = liveValueMax - db.Max;


            double min = signChange == 1 ? pm.MasterValue - offsetMin : pm.MasterValue + offsetMin;
            double max = signChange == 1 ? pm.MasterValue - offsetMax : pm.MasterValue + offsetMax;

            if (mode == ProcedureMode.Measurement && compensation != 0)
            {
                min += compensation;
                max += compensation;
            }

            return (Math.Round(min, 3), Math.Round(max, 3));
        }



        private (double Min, double Max) CalculateID1(
     Dictionary<string, ProbeMeasurement> probeMeasurements,
     Dictionary<string, (double Min, double Max)> dbRefDict,
     ProcedureMode mode,
     int signChange,
     double compensation,
     int activeIdValue)
        {
            if (!probeMeasurements.TryGetValue("ID-1", out var pm))
                return (double.NaN, double.NaN);

            double liveValueMax = pm.Readings.Any() ? pm.MaxValue : 0;
            double liveValueMin = pm.Readings.Any() ? pm.MinValue : 0;
            if (liveValueMax == 0)
                return (double.NaN, double.NaN);

            if (!dbRefDict.TryGetValue("ID-1", out var db))
                return (double.NaN, double.NaN);

            double offsetMin = liveValueMin - db.Min;
            double offsetMax = liveValueMax - db.Max;

            double min = signChange == 1 ? pm.MasterValue - offsetMin : pm.MasterValue + offsetMin;
            double max = signChange == 1 ? pm.MasterValue - offsetMax : pm.MasterValue + offsetMax;

            if (mode == ProcedureMode.Measurement && compensation != 0)
            {
                min += compensation;
                max += compensation;
            }

            return (Math.Round(min, 3), Math.Round(max, 3));
        }


        private (double Min, double Max) CalculateID2(
     Dictionary<string, ProbeMeasurement> probeMeasurements,
     Dictionary<string, (double Min, double Max)> dbRefDict,
     ProcedureMode mode,
     int signChange,
     double compensation,
     int activeIdValue)
        {
            if (!probeMeasurements.TryGetValue("ID-2", out var pm))
                return (double.NaN, double.NaN);

            double liveValueMax = pm.Readings.Any() ? pm.MaxValue : 0;
            double liveValueMin = pm.Readings.Any() ? pm.MinValue : 0;
            if (liveValueMax == 0)
                return (double.NaN, double.NaN);

            if (!dbRefDict.TryGetValue("ID-2", out var db))
                return (double.NaN, double.NaN);

            double offsetMin = liveValueMin - db.Min;
            double offsetMax = liveValueMax - db.Max;


            double min = signChange == 1 ? pm.MasterValue - offsetMin : pm.MasterValue + offsetMin;
            double max = signChange == 1 ? pm.MasterValue - offsetMax : pm.MasterValue + offsetMax;

            if (mode == ProcedureMode.Measurement && compensation != 0)
            {
                min += compensation;
                max += compensation;
            }

            return (Math.Round(min, 3), Math.Round(max, 3));
        }


        private double CalculateOverallLength(
    Dictionary<string, ProbeMeasurement> probeMeasurements,
    Dictionary<string, (double Min, double Max)> dbRefDict,
    ProcedureMode mode,
    int signChange,
    double compensation)
        {
            if (!probeMeasurements.TryGetValue("OL", out var pm))
                return double.NaN;

            double liveValue = pm.Readings.Any() ? pm.MaxValue : 0;
            if (liveValue == 0)
                return double.NaN;

            if (!dbRefDict.TryGetValue("OL", out var db))
                return double.NaN;

            // ✅ Use single reference value (typically db.Max or average)
            double dbRefValue = db.Max;  // Master reference value
            double offset = liveValue - dbRefValue;

            // ✅ Single calculated value
            double value = signChange == 1 ? pm.MasterValue - offset : pm.MasterValue + offset;

            if (mode == ProcedureMode.Measurement && compensation != 0)
                value += compensation;

            return Math.Round(value, 3);  // ✅ Single value only
        }



        #endregion

        #region 🔥 RUNOUT PARAMETERS (Max - Min, uses individual probes)

        private double  CalculateStepRunout1(
      Dictionary<string, ProbeMeasurement> probeMeasurements,
      Dictionary<string, (double Min, double Max)> dbRefDict,
      ProcedureMode mode,
      int signChange,
      double compensation)
        {
            return CalculateRunoutCommon("RN1", probeMeasurements, mode, signChange, compensation);
        }


        private double CalculateStepRunout2(
    Dictionary<string, ProbeMeasurement> probeMeasurements,
    Dictionary<string, (double Min, double Max)> dbRefDict,
    ProcedureMode mode,
    int signChange,
    double compensation)
        {
            return CalculateRunoutCommon("RN2", probeMeasurements, mode, signChange, compensation);
        }


        private double CalculateRN1(
    Dictionary<string, ProbeMeasurement> probeMeasurements,
    Dictionary<string, (double Min, double Max)> dbRefDict,
    ProcedureMode mode,
    int signChange,
    double compensation)
        {
            return CalculateRunoutCommon("RN3", probeMeasurements, mode, signChange, compensation);
        }


        private double CalculateRN2(
    Dictionary<string, ProbeMeasurement> probeMeasurements,
    Dictionary<string, (double Min, double Max)> dbRefDict,
    ProcedureMode mode,
    int signChange,
    double compensation)
        {
            return CalculateRunoutCommon("RN4", probeMeasurements, mode, signChange, compensation);
        }


        private double CalculateRN3(
    Dictionary<string, ProbeMeasurement> probeMeasurements,
    Dictionary<string, (double Min, double Max)> dbRefDict,
    ProcedureMode mode,
    int signChange,
    double compensation)
        {
            return CalculateRunoutCommon("RN5", probeMeasurements, mode, signChange, compensation);
        }


        private double CalculateRN4(
    Dictionary<string, ProbeMeasurement> probeMeasurements,
    Dictionary<string, (double Min, double Max)> dbRefDict,
    ProcedureMode mode,
    int signChange,
    double compensation)
        {
            return CalculateRunoutCommon("RN6", probeMeasurements, mode, signChange, compensation);
        }

        private double CalculateRN5(
    Dictionary<string, ProbeMeasurement> probeMeasurements,
    Dictionary<string, (double Min, double Max)> dbRefDict,
    ProcedureMode mode,
    int signChange,
    double compensation)
        {
            return CalculateRunoutCommon("RN7", probeMeasurements, mode, signChange, compensation);
        }


        private double CalculateRunoutCommon(
    string key,
    Dictionary<string, ProbeMeasurement> probeMeasurements,
    ProcedureMode mode,
    int signChange,
    double compensation)
        {
            if (!probeMeasurements.TryGetValue(key, out var pm))
                return double.NaN;

            if (!pm.Readings.Any())
                return double.NaN;

            // ✅ CORRECT: Runout = Max - Min of LIVE readings only
            double runout = Math.Abs(pm.MaxValue - pm.MinValue);
            runout = Math.Round(runout, 3);

            // Apply compensation for Measurement mode
            if (mode == ProcedureMode.Measurement && compensation != 0)
                runout += compensation;

            return Math.Abs(runout);  // ✅ Single runout value
        }


        #endregion









        private void LoadProbeConfigurations(string partCode)
        {
            _probeMeasurements.Clear();
            _orderedProbeMeasurements.Clear();

            var probeInstalls = _dataStorageService.GetProbeInstallByPartNumber(partCode);
            var masterVals = _dataStorageService.GetPartConfig(partCode);

            foreach (var probe in probeInstalls)
            {
                // 🔹 Use probe.ParameterName to match master
                string parameterKey = probe.ParameterName?.Trim();

                var config = masterVals?
                    .Where(m =>
                        string.Equals(m.Parameter?.Trim(), parameterKey, StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(m.Para_No?.Trim(), parameterKey, StringComparison.OrdinalIgnoreCase)
                    )
                    .LastOrDefault(); // last master value if multiple matches exist

                double masterVal = config?.Nominal ?? 0;
                double tolPlus = config?.RTolPlus ?? 0;
                double tolMinus = config?.RTolMinus ?? 0;

                // 🔹 Make unique key per probe
                var uniqueId = probe.ProbeName;

                var pm = new ProbeMeasurement
                {
                    ProbeId = uniqueId,
                    Name = probe.ParameterName,
                    Readings = new List<double>(),
                    MasterValue = masterVal,
                    TolerancePlus = tolPlus,
                    ToleranceMinus = tolMinus
                };

                _probeMeasurements[uniqueId] = pm;
                _orderedProbeMeasurements.Add(pm);

                // 🔍 Debug
                foreach (var pms in _orderedProbeMeasurements)
                {
                    System.Diagnostics.Debug.WriteLine($"  Probe {pm.ProbeId} ({pm.Name}): Master={pm.MasterValue}, Tol±={pm.TolerancePlus}/{pm.ToleranceMinus}");
                }

            }
        }



        private async void LoadProbeConfigurationsforMasterInspection(string partCode)
        {
            _probeMeasurements.Clear();
            _orderedProbeMeasurements.Clear();

            var probeInstalls = _dataStorageService.GetProbeInstallByPartNumber(partCode);
            var masterVals = _dataStorageService.GetPartConfig(partCode);

            foreach (var probe in probeInstalls)
            {
                // 🔹 Use probe.ParameterName to match master
                string parameterKey = probe.ParameterName?.Trim();

                var config = masterVals?
                    .Where(m =>
                        string.Equals(m.Parameter?.Trim(), parameterKey, StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(m.Para_No?.Trim(), parameterKey, StringComparison.OrdinalIgnoreCase)
                    )
                    .LastOrDefault(); // last master value if multiple matches exist


                double masterVal = config?.Nominal ?? 0;
                double tolPlus = config?.RTolPlus ?? 0;
                double tolMinus = config?.RTolMinus ?? 0;
                int Sign = config?.Sign_Change ?? 0;
                double Comp = config?.Compensation ?? 0;

                // ✅ Use ProbeName as unique ID (no channel concatenation)
                var uniqueId = probe.ProbeName;

                var pm = new ProbeMeasurement
                {
                    ProbeId = uniqueId,                  // ProbeName (unique from DB)
                    Name = probe.ParameterName,         // Human readable display name
                    Readings = new List<double>(),
                    MasterValue = masterVal,
                    TolerancePlus = tolPlus,
                    ToleranceMinus = tolMinus,
                    SignChange = Sign,
                    Compensation = Comp,
                };

                _probeMeasurements[uniqueId] = pm;
                _orderedProbeMeasurements.Add(pm);
            }

            //System.Diagnostics.Debug.WriteLine($"[DEBUG] Loaded {_probeMeasurements.Count} probe configs for part {partCode}");
            foreach (var pm in _orderedProbeMeasurements)
            {
                System.Diagnostics.Debug.WriteLine($"  Probe {pm.ProbeId} ({pm.Name}): Master={pm.MasterValue}, Tol±={pm.TolerancePlus}/{pm.ToleranceMinus}");
            }

            //await StartMeasurementProcessAsync();
        }



        private async Task StartMeasurementProcessAsync()
        {
            try
            {
                NotifyStatus("Initializing measurement...");

                // 1️⃣ Connect PLC + Serial
                //if (!_plcProbeService.IsConnected)
                //    await _plcProbeService.ConnectAsync();

                // 2️⃣ Load probes for this part
                await _plcProbeService.LoadProbesAsync(_currentPartCode);

                // ❗ Ensure probes loaded
                if (_plcProbeService.ProbeReadings.Count == 0)
                {
                    MessageBox.Show("No probes found for this Part No.", "Error");
                    return;
                }

                // 3️⃣ Prepare measurement list
                _orderedProbeMeasurements = _probeMeasurements.Values.ToList();

                // 4️⃣ Call your main function
                await RunMotorAndCollectReadingsAsync(_orderedProbeMeasurements, ProcedureMode.Measurement);

                NotifyStatus("Measurement complete.");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error");
            }
        }




        public void Dispose()
        {
            Cleanup();
            _dataStorageService?.Dispose();

            GC.SuppressFinalize(this);  // ✅ Suppress finalizer
        }


        public class MasterCompletedEventArgs : EventArgs
        {
            public Dictionary<string, double> MasteredValues { get; }
            public bool Success { get; }

            public MasterCompletedEventArgs(Dictionary<string, double> values, bool success)
            {
                MasteredValues = values;
                Success = success;
            }
        }

        private void OnMasterCompleted(Dictionary<string, double> values, bool success)
        {
            MasterCompleted?.Invoke(this, new MasterCompletedEventArgs(values, success));
        }
        // For parameters that have 2 probes: build a virtual probe = max of both
        private Dictionary<string, ProbeMeasurement> BuildEffectiveProbesByName(
            List<ProbeMeasurement> probes)
        {
            // group by Name (STEP OD 1, OD-1, etc.)
            var result = new Dictionary<string, ProbeMeasurement>();

            foreach (var group in probes.Where(p => !string.IsNullOrEmpty(p.Name))
                                        .GroupBy(p => p.Name))
            {
                var list = group.ToList();

                // only one probe for this parameter -> use as is
                if (list.Count == 1)
                {
                    var single = list[0];
                    result[group.Key] = single;
                    continue;
                }

                // two probes (or more) -> element-wise max
                var p1 = list[0];
                var p2 = list[1];

                int count = Math.Min(p1.Readings.Count, p2.Readings.Count);
                var merged = new ProbeMeasurement
                {
                    ProbeId = $"{p1.Name}_MERGED",
                    Name = p1.Name,
                    Readings = new List<double>(count),
                    MasterValue = p1.MasterValue,       // same master for that parameter
                    TolerancePlus = p1.TolerancePlus,
                    ToleranceMinus = p1.ToleranceMinus,
                    SignChange = p1.SignChange,
                    Compensation = p1.Compensation
                };

                for (int i = 0; i < count; i++)
                {
                    double v = Math.Max(p1.Readings[i], p2.Readings[i]);
                    merged.Readings.Add(v);
                }

                var clean = merged.Readings.Where(x => !double.IsNaN(x)).ToList();
                merged.MaxValue = clean.Count > 0 ? clean.Max() : 0;
                merged.MinValue = clean.Count > 0 ? clean.Min() : 0;

                result[group.Key] = merged;
            }

            return result;
        }



        private void NotifyStatus(string message)
        {
            MainWindow.ShowStatusMessage(message);
        }




        private async Task NotifyOnUIAsync(string message)
        {
            if (Application.Current?.Dispatcher?.CheckAccess() == true)
                NotifyStatus(message);
            else
                await Application.Current.Dispatcher.InvokeAsync(() => NotifyStatus(message));
        }


        public class ProbeReadingEventArgs : EventArgs
        {
            public string ModuleId { get; set; } = "";
            public double Value { get; set; }

            public ProbeReadingEventArgs(string moduleId, double value)
            {
                ModuleId = moduleId;
                Value = value;
            }
        }



        private bool ShouldContinueMeasurement()
        {
            return _continueMeasurement;
        }



        // In MasterService class
        public void ResetRejectionBits()
        {
            SetPlcDevice("M302", 0); // General rejection
            SetPlcDevice("M303", 0); // SRO rejection
            SetPlcDevice("M304", 0); // STDIA rejection
            SetPlcDevice("M305", 0); // Seat Height rejection
            SetPlcDevice("M307", 0); // Groove Diameter/Position rejection
        }




        public class ParameterInfo
        {
            public string Name { get; set; }
            public int SignChange { get; set; }
            public double Compensation { get; set; }
        }


    }

}
