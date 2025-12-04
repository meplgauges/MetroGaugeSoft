using ActUtlType64Lib;
using DocumentFormat.OpenXml.EMMA;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;
using Windows.Storage.Streams;
using Windows.UI.WindowManagement;
using static SkiaSharp.HarfBuzz.SKShaper;

namespace EVMS.Service
{
    public class ProbeMeasurement
    {
        public string? ProbeId { get; set; }
        public string? Name { get; set; }
        public List<double> Readings { get; set; } = new List<double>();
        public double MaxValue { get; set; } = 0;
        public double MinValue { get; set; } = 0;  // Add this property
        public double MasterValue { get; set; } = 0;
        public double TolerancePlus { get; set; }
        public double ToleranceMinus { get; set; }
        public int SignChange { get; set; }        // 0 = add, 1 = subtract
        public double Compensation { get; set; }
    }


    public class ParameterResult
    {
        public double Value { get; set; }
        public bool IsOk { get; set; }
    }
 

    public class MasterService : IDisposable
    {
        public event Action? MeasurementStarted;
        public event Action? MeasurementStopped;

        public bool _continueMeasurement = false;

        public bool IsMeasurementRunning { get; set; } = false;



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
            _plcProbeService.ProbeValueUpdated += ProbeValueUpdatedHandler;
            ArraySize = _dataStorageService.GetReadingCount();

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

        public async Task<bool> EnsureConnectionAsync()
        {
            if (_plcProbeService == null) return false;

            if (!_plcProbeService.IsConnected)
            {
                bool connected = await _plcProbeService.ConnectAsync();
                if (connected)
                {
                    _plcProbeService.ProbeValueUpdated += ProbeValueUpdatedHandler;
                    int openResult = plc.Open();
                    if (openResult != 0)
                    {
                        MessageBox.Show($"Failed to open PLC connection. Error code: {openResult}", "PLC Init Error", MessageBoxButton.OK, MessageBoxImage.Error);
                        return false;
                    }
                }
                return connected;
            }
            return true;
        }

        private void ProbeValueUpdatedHandler(object? sender, ProbeReadingEventArgs e)
        {
            _collectedReadings.Enqueue((e.ModuleId, e.Value));
        }

        public bool SetPlcDevice(string device, int value)
        {
            int ret = plc.SetDevice(device, (short)value);
            if (ret != 0)
            {
                NotifyStatus("SetDevice failed!!");

                return false;
            }
            return true;
        }

        public int GetPlcDeviceBit(string device)
        {
            if (plc.GetDevice(device, out int value) != 0)
            {
                NotifyStatus($"GetDevice failed for bit {device}! Exception");
                return -1; // Indicate failure
            }
            return value; // Return the bit value read from PLC device
        }


        private void Cleanup()
        {
            _plcProbeService.StopLiveReading();
            _plcProbeService?.Disconnect();
            _plcProbeService.ProbeValueUpdated -= ProbeValueUpdatedHandler;
            plc.Close();
        }

        public enum ProcedureMode
        {
            Mastering,
            MasterInspection,
            Measurement
        }

        public async Task MasterCheckProcedureAsync(ProcedureMode mode)
        {
            try
            {
                var autoList = _dataStorageService.GetActiveBit();
                if (autoList == null || autoList.Count == 0)
                {
                    MessageBox.Show("No Settings Found.", "Info", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                var autoControl = autoList.FirstOrDefault(c => string.Equals(c.Description, "Auto/Manual", StringComparison.OrdinalIgnoreCase));
                int bitValue = GetPlcDeviceBit("X14");

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

                if (autoControl?.Bit == 1)
                {
                    SetPlcDevice("M400", 1);
                    SetPlcDevice("M100", 0);
                    SetPlcDevice("M300", 0);
                    switch (mode)
                    {
                        case ProcedureMode.Mastering:
                            SetPlcDevice("M100", 1);
                            break;
                        case ProcedureMode.MasterInspection:
                            SetPlcDevice("M100", 1);
                            break;
                        case ProcedureMode.Measurement:
                            SetPlcDevice("M300", 1);
                            //LoadProbeConfigurations(_currentPartCode);

                            break;
                    }
                }

                SetPlcDevice("M10", 0);
                if (mode==ProcedureMode.MasterInspection)
                {
                    LoadProbeConfigurationsforMasterInspection(_currentPartCode);

                }
                else
                {
                    LoadProbeConfigurations(_currentPartCode);

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
                    // Optionally handle any post-measurement processing her
                    return;
                }


                bool firstMeasurementCycle = true;
                if (mode == ProcedureMode.Mastering || mode == ProcedureMode.MasterInspection)
                {
                    if (autoControl?.Bit == 0)
                    {

                        string loadMsg = mode == ProcedureMode.Mastering ? "Load The Value in fixture..." : "Place part for measurement...";
                        await NotifyOnUIAsync(loadMsg);
                        await WaitForValidProbeReadingAsync("100AY33P34");
                        string promptMsg = mode == ProcedureMode.Mastering ? "PRESS START SWITCH TO START MASTERING" : "PRESS START SWITCH TO START MasterInspection";
                        await NotifyOnUIAsync(promptMsg);
                        while (GetPlcDeviceBit("X1") != 1) await Task.Delay(100);
                        await NotifyOnUIAsync("START BUTTON PRESSED");
                    }
                    else
                    {
                        
                        string startMsg = mode == ProcedureMode.Mastering ? "Press Robo start button to start mastering." : "Press Robo start button to start MasterInspection.";
                        await NotifyOnUIAsync(startMsg);
                        while (GetPlcDeviceBit("M20") != 1) await Task.Delay(100);
                        await NotifyOnUIAsync("Robo start button Pressed");//DIRECTLY GETTING THE ROBO START
                        SetPlcDevice("M101", 1);
                        await NotifyOnUIAsync("Waiting Robot to Load Part...");
                        await WaitForValidProbeReadingAsync("100AY33P34");
                        await NotifyOnUIAsync("Waiting Robot for Safe Position...");
                        while (GetPlcDeviceBit("M101") != 0) await Task.Delay(100);
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

        // ** Separate measurement function for external toggle to run cycle **
        //public async Task RunMeasurementCycleAsync()
        //{
        //    var autoList = _dataStorageService.GetActiveBit();
        //    var autoControl = autoList?.FirstOrDefault(c =>
        //        string.Equals(c.Description, "Auto/Manual", StringComparison.OrdinalIgnoreCase));
        //    int bitValue = GetPlcDeviceBit("X14"); // Auto/Manual PLC bit

        //    if (!(autoControl != null && (autoControl.Bit == bitValue)))
        //    {
        //        if (autoControl?.Bit == 0 && bitValue == 1)
        //            await NotifyOnUIAsync("Software is in Manual mode. Please switch PLC to Manual mode.");
        //        else if (autoControl?.Bit == 1 && bitValue == 0)
        //            await NotifyOnUIAsync("Software is in Auto mode. Change PLC to Auto mode.");
        //        return;
        //    }

        //    var sortedProbeMeasurements = _orderedProbeMeasurements;
        //    bool firstMeasurementCycle = true;

        //    do
        //    {
        //        // Only check _continueMeasurement before starting a new cycle
        //        if (!firstMeasurementCycle && !_continueMeasurement)
        //        {
        //            await NotifyOnUIAsync("Stop requested. Finishing current cycle before stopping...");
        //            break; // exit loop after finishing current cycle
        //        }

        //        // Clear previous readings if not the first cycle
        //        if (!firstMeasurementCycle)
        //        {
        //            foreach (var probe in sortedProbeMeasurements)
        //                probe.Readings.Clear();
        //        }

        //        // Wait for robot safe position if not the first cycle
        //        if (!firstMeasurementCycle)
        //        {
        //            await NotifyOnUIAsync("Waiting for robot to reach safe position...");
        //            while (GetPlcDeviceBit("M301") != 0) await Task.Delay(100);
        //            await NotifyOnUIAsync("Robot is in safe position. Ready to load next part.");
        //        }

        //        // ===== Start Cycle =====
        //        if (firstMeasurementCycle)
        //        {
        //            if (autoControl?.Bit == 0)
        //            {
        //                //await NotifyOnUIAsync("Place part for measurement...");
        //                //await WaitForValidProbeReadingAsync("100AY08P42");

        //                //await NotifyOnUIAsync("PRESS START SWITCH TO START MEASUREMENT");
        //                //while (GetPlcDeviceBit("X1") != 1) await Task.Delay(100);

        //                //await NotifyOnUIAsync("START BUTTON PRESSED");
        //            }
        //            else
        //            {
        //                if (!_continueMeasurement) break;

        //                await NotifyOnUIAsync("Press Robo start button to start measurement.");
        //                while (GetPlcDeviceBit("M20") != 1 && _continueMeasurement) await Task.Delay(100);

        //                if (!_continueMeasurement) break;

        //                await NotifyOnUIAsync("Robo start button Pressed");
        //                SetPlcDevice("M301", 1);

        //                await NotifyOnUIAsync("Waiting Robot to Load Part...");
        //                if (_continueMeasurement)
        //                    await WaitForValidProbeReadingAsync("100AY33P34");

        //                if (!_continueMeasurement) break;

        //                await NotifyOnUIAsync("Waiting Robot for Safe Position...");
        //                while (GetPlcDeviceBit("M301") != 0 && _continueMeasurement) await Task.Delay(100);
        //            }
        //        }
        //        else
        //        {
        //            if (!_continueMeasurement) break;

        //            await NotifyOnUIAsync("Robo start button Pressed");
        //            SetPlcDevice("M301", 1);

        //            await NotifyOnUIAsync("Waiting Robot to Load Part...");
        //            if (_continueMeasurement)
        //                await WaitForValidProbeReadingAsync("100AY33P34");

        //            if (!_continueMeasurement) break;

        //            await NotifyOnUIAsync("Waiting Robot for Safe Position...");
        //            while (GetPlcDeviceBit("M301") != 0 && _continueMeasurement) await Task.Delay(100);
        //        }

        //        /// Only run motor if measurement is ON
        //        if (_continueMeasurement)
        //        {
        //            await RunMotorAndCollectReadingsAsync(sortedProbeMeasurements, ProcedureMode.Measurement);
        //        }


        //        var probeMeasurementByNameMeasurement = sortedProbeMeasurements
        //            .Where(pm => !string.IsNullOrEmpty(pm.Name))
        //            .ToDictionary(pm => pm.Name!);

        //        HandleMeasurementStage(probeMeasurementByNameMeasurement, _currentPartCode);

        //        var readingsSummary = string.Join("; ",
        //            sortedProbeMeasurements.Select(p =>
        //                $"{p.Name}: {string.Join(",", p.Readings)}"));

        //        await NotifyOnUIAsync($"Cycle completed. Probe readings: {readingsSummary}");

        //        firstMeasurementCycle = false;

        //        await NotifyOnUIAsync("Ready for next cycle.");

        //    } while (ShouldContinueMeasurement());

        //    // ===== Stop & Reset =====
        //    await NotifyOnUIAsync("Measurement cycle stopped by user.");

        //    SetPlcDevice("M10", 0);   // Motor OFF
        //                              //SetPlcDevice("M20", 0);   // Robo Start OFF if needed
        //    SetPlcDevice("M300", 0);  // Measurement Complete OFF
        //                              //SetPlcDevice("M301", 0);  // Robot Safe Position Reset if needed

        //    MeasurementStopped?.Invoke();
        //}

        // Helper should continue method

        public async Task RunMeasurementCycleAsync()
        {
            var autoList = _dataStorageService.GetActiveBit();
            var autoControl = autoList?.FirstOrDefault(c =>
                string.Equals(c.Description, "Auto/Manual", StringComparison.OrdinalIgnoreCase));
            int bitValue = GetPlcDeviceBit("X14"); // Auto/Manual PLC bit

            if (!(autoControl != null && (autoControl.Bit == bitValue)))
            {
                if (autoControl?.Bit == 0 && bitValue == 1)
                    await NotifyOnUIAsync("Software is in Manual mode. Please switch PLC to Manual mode.");
                else if (autoControl?.Bit == 1 && bitValue == 0)
                    await NotifyOnUIAsync("Software is in Auto mode. Change PLC to Auto mode.");
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
                    await NotifyOnUIAsync("Waiting for Robot to reach safe position...");
                    while (GetPlcDeviceBit("M301") != 0) await Task.Delay(100);
                    await NotifyOnUIAsync("Robot is in safe position. Ready to load next part.");
                }

                // ===== Start Cycle =====
                if (firstMeasurementCycle)
                {
                    if (autoControl?.Bit == 0)
                    {
                       

                        await WaitForValidProbeReadingAsync("100AY33P34");

                        if (!_continueMeasurement) break;

                        await NotifyOnUIAsync("Part detected. Press the Start switch to begin measurement");
                        while (GetPlcDeviceBit("X1") != 1 && _continueMeasurement) await Task.Delay(100);

                        if (!_continueMeasurement) break;

                        await NotifyOnUIAsync("Starting measurement...");

                    }
                    else
                    {
                        // === AUTO MODE ===
                        if (!_continueMeasurement) break;

                        await NotifyOnUIAsync("Press the Robo Start button to begin measurement..");
                        while (GetPlcDeviceBit("M20") != 1 && _continueMeasurement) await Task.Delay(100);

                        if (!_continueMeasurement) break;

                        await NotifyOnUIAsync("Robo Start button pressed");
                        SetPlcDevice("M301", 1);

                        await NotifyOnUIAsync("Waiting for Robot to Load Part...");
                        if (_continueMeasurement)
                            await WaitForValidProbeReadingAsync("100AY33P34");

                        if (!_continueMeasurement) break;

                   
                        await NotifyOnUIAsync("Waiting for Robot to Safe Position...");
                        while (GetPlcDeviceBit("M301") != 0 && _continueMeasurement) await Task.Delay(100);
                    }
                }
                else
                {
                    if (autoControl?.Bit == 0)
                    {
                        //await NotifyOnUIAsync(" Remove the Part");

                        await WaitForPartRemovedAsync("100AY33P34");
                        await Task.Delay(1000); // stabilization delay

                        await NotifyOnUIAsync("Part removed..");

                            if (!_continueMeasurement)
                                break;

                            // 1️⃣ Ask operator to load the part
                            await NotifyOnUIAsync("Please load part...");
                            await WaitForValidProbeReadingAsync("100AY33P34");

                            if (!_continueMeasurement)
                                break;

                            // 2️⃣ Ask operator to press start switch
                            await NotifyOnUIAsync("Part detected. Press Start switch to begin measurement.");
                            while (GetPlcDeviceBit("X1") != 1 && _continueMeasurement)
                                await Task.Delay(100);

                            if (!_continueMeasurement)
                                break;

                            await NotifyOnUIAsync("Starting measurement..");

                          
                            // 6️⃣ Wait for part to be removed
                           
                            // 7️⃣ Small delay before next cycle
                            //await Task.Delay(1000);

                        } 

                    else
                    {
                        // === AUTO MODE REPEAT ===
                        if (!_continueMeasurement) break;

                        await NotifyOnUIAsync("Robo Start button pressed");
                        SetPlcDevice("M301", 1);

                        await NotifyOnUIAsync("Waiting for Robot to load part..");
                        if (_continueMeasurement)
                            await WaitForValidProbeReadingAsync("100AY33P34");

                        if (!_continueMeasurement) break;

                        await NotifyOnUIAsync("Waiting for Robot to reach safe position....");
                        while (GetPlcDeviceBit("M301") != 0 && _continueMeasurement) await Task.Delay(100);
                    }
                }

                // 🛠 Motor should NOT run again after measurement in manual mode
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

            SetPlcDevice("M10", 0);   // Motor OFF
            SetPlcDevice("M300", 0);  // Measurement Complete OFF
            MeasurementStopped?.Invoke();
        }


        public async Task WaitForValidProbeReadingAsync(string targetProbeId, bool suppressDetectedMessage = false)
        {
            await NotifyOnUIAsync("Initializing probe readings...");

            _plcProbeService.StartLiveReading(100);

            bool partDetected = false;
            bool messageFired = false;

            if (!suppressDetectedMessage)
                if (GetPlcDeviceBit("X14") == 1)
                    await NotifyOnUIAsync("Waiting the robo to load part");
                else
                    await NotifyOnUIAsync("Load the part");
            while (!partDetected)
            {
                await Task.Delay(100);

                var probeVal = _collectedReadings
                    .Where(r => r.ProbeId == targetProbeId)
                    .Select(r => r.Value)
                    .LastOrDefault();

                bool partPresent = Math.Abs(probeVal) > 0.100;

                if (partPresent && !messageFired)
                {
                    messageFired = true;
                    partDetected = true;
                    if (!suppressDetectedMessage)
                        await NotifyOnUIAsync("Part detected. Proceeding...");
                }
                else if (!partPresent && messageFired)
                {
                    messageFired = false;
                    if (!suppressDetectedMessage)
                        await NotifyOnUIAsync("Part removed. Waiting for new part...");
                }
                else if (!partPresent && !messageFired)
                {
                    if (!suppressDetectedMessage)
                        if (GetPlcDeviceBit("X14") == 1)
                            await NotifyOnUIAsync("Waiting the robo to load part");
                        else
                            await NotifyOnUIAsync("Load the part");
                }
            }

            _plcProbeService.StopLiveReading();

            if (!suppressDetectedMessage)
                await NotifyOnUIAsync("Values updated...");
        }


        // 🆕 Added Helper for Manual Mode (wait until part removed)
        public async Task WaitForPartRemovedAsync(string targetProbeId, bool suppressRemovedMessage = false)
        {
            await NotifyOnUIAsync("Remove The Part...");

            _plcProbeService.StartLiveReading(100);

            bool partRemoved = false;

            while (!partRemoved && _continueMeasurement)
            {
                await Task.Delay(100);

                var probeVal = _collectedReadings
                    .Where(r => r.ProbeId == targetProbeId)
                    .Select(r => r.Value)
                    .LastOrDefault();

                bool partPresent = Math.Abs(probeVal) > 0.100;

                if (!partPresent)
                {
                    partRemoved = true;
                    if (!suppressRemovedMessage)
                        await NotifyOnUIAsync("Part removed. Ready for next part.");
                }
            }

            _plcProbeService.StopLiveReading();

            //if (!suppressRemovedMessage)
            //    await NotifyOnUIAsync("Waiting for next cycle...");
        }




        private async Task RunMotorAndCollectReadingsAsync(List<ProbeMeasurement> sortedProbeMeasurements, ProcedureMode mode)
        {
            
            
            SetPlcDevice("M10", 1);
            NotifyStatus("Collecting the readings...");

            // Wait for PLC ready signal
            while (GetPlcDeviceBit("X0") != 1)
                await Task.Delay(100);

            SetPlcDevice("M14", 1);

            // Clear old data once
            while (_collectedReadings.TryDequeue(out _)) { }

            // Start live reading once
            _plcProbeService.StartLiveReading();

            for (int i = 0; i < ArraySize; i++)
            {
                if (Abort)
                {
                    MessageBox.Show("Operation aborted.", "Abort", MessageBoxButton.OK, MessageBoxImage.Warning);
                    break;
                }

                await Task.Delay(15);

                while (_collectedReadings.TryDequeue(out var reading))
                {
                    if (_probeMeasurements.TryGetValue(reading.ProbeId, out var pm))
                    {
                        if (pm.Readings.Count < ArraySize)
                        {
                            pm.Readings.Add(reading.Value);
                        }
                    }
                }

                if (_probeMeasurements.Values.All(pm => pm.Readings.Count >= ArraySize))
                {
                    // All probes have enough readings, stop the loop
                    break;
                }
            }

            // Stop reading
            _plcProbeService.StopLiveReading();
            SetPlcDevice("M14", 0);
            SetPlcDevice("M10", 0);

            // Process readings for each probe
            foreach (var pm in sortedProbeMeasurements)
            {
                // Clean up readings
                var validReadings = pm.Readings
                    .Where(r => !double.IsNaN(r))
                    .OrderBy(r => r) // sort values (ascending)
                    .ToList();

                if (validReadings.Count > 0)
                {
                    pm.MaxValue = validReadings.Max();
                    pm.MinValue = validReadings.Min(); // always calculate min
                }
                else
                {
                    pm.MaxValue = 0;
                    pm.MinValue = 0;
                }

            }

            NotifyStatus("Reading collection completed.");
        }

        private async Task HandleProcedurePostProcessingAsync(ProcedureMode mode, List<ProbeMeasurement> probeMeasurements, bool isFirstMeasurementCycle)
        {
            switch (mode)
            {
                case ProcedureMode.Mastering:
                    await NotifyOnUIAsync("Mastering Completed. Press Enter to Inspect the Master");

                    var masterValues = probeMeasurements
                        .Where(pm => !string.IsNullOrEmpty(pm.Name))
                        .ToDictionary(pm => pm.Name!, pm => pm.MaxValue);

                    OnCalculatedValuesReady(masterValues);

                    await HandleMasteringStageAsync();
                    break;

                case ProcedureMode.MasterInspection:
                    {
                        // Build dictionary safely (avoid duplicate keys)
                        var probeMeasurementByName = probeMeasurements
                            .Where(pm => !string.IsNullOrEmpty(pm.Name))
                            .GroupBy(pm => pm.Name)
                            .ToDictionary(g => g.Key, g => g.First());

                        // Perform master inspection
                        HandleMasterInspectionStage(probeMeasurementByName, _currentPartCode);

                        // 🔹 Read Auto/Manual PLC Bit (X14)
                        int bitValue = GetPlcDeviceBit("X14");  // 1 = Auto mode

                        if (bitValue == 1)
                        {
                            // 🔸 Trigger Master Inspection signal
                            SetPlcDevice("M102", 1);

                            await NotifyOnUIAsync("Waiting Safe position from Robot...");

                            // 🔸 Wait until robot finishes (M102 reset by PLC)
                            while (GetPlcDeviceBit("M102") != 0)
                                await Task.Delay(100);

                            await NotifyOnUIAsync("Master Inspection Completed ✅");
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

            foreach (var pm in _probeMeasurements.Values)
            {
                _dataStorageService.SaveProbeReadings(
                    _dataStorageService.GetProbeInstallByPartNumber(_currentPartCode),
                    _currentPartCode,
                    new Dictionary<string, double> { { pm.ProbeId, pm.MaxValue } }
                );
            }

            MasterComplete = true;
            var resultsWithStatus = _probeMeasurements.Values
                .Where(pm => !string.IsNullOrEmpty(pm.Name))
                .ToDictionary(
                    pm => pm.Name!,
                    pm => new ParameterResult { Value = pm.MaxValue, IsOk = true } // assuming OK on mastering
                );
            OnCalculatedValuesWithStatusReady(resultsWithStatus);

            // Determine current mode
            var autoList = _dataStorageService.GetActiveBit();
            var autoControl = autoList?.FirstOrDefault(c => string.Equals(c.Description, "Auto/Manual", StringComparison.OrdinalIgnoreCase));
            int bitValue = GetPlcDeviceBit("X14"); // PLC Auto/Manual bit

            bool isAuto = autoControl != null && autoControl.Bit == 1 && bitValue == 1;

            if (isAuto)
            {
                SetPlcDevice("M102", 1); // trigger robot to move to safe position
                await NotifyOnUIAsync("Waiting Safe position from Robot...");
                while (GetPlcDeviceBit("M102") != 0)
                    await Task.Delay(100);
            }

            // Mastering complete message
            await NotifyOnUIAsync("Mastering completed. Press Enter to inspect the master.");
        }


        private bool IsValidParameter(string param, Dictionary<string, ProbeMeasurement> probeMeasurements, Dictionary<string, double> dbRefDict)
        {
            return probeMeasurements.ContainsKey(param) && dbRefDict.ContainsKey(param);
        }

        private void HandleMasterInspectionStage(Dictionary<string, ProbeMeasurement> probeMeasurements, string partCode)
        {
            var dbRefList = _dataStorageService.GetMasterProbeRef(_currentPartCode);

            var mode = ProcedureMode.MasterInspection;

            // ✅ FIXED: Prevent "duplicate key" crash
            var dbRefDict = dbRefList
                .GroupBy(x => x.Name)
                .ToDictionary(g => g.Key, g => g.First().Value);

            // Fetch master part configurations including tolerances
            var masterVals = _dataStorageService.GetMasterReadingByPart(partCode);
           // var masterp = _dataStorageService.GetPartConfig(partCode);

            // 17 measurement parameters in order

            var parameterNames = masterVals
                .Select(m => m.Parameter)
                .Distinct()
                .ToList();

            //    string[] parameterNames = new string[]
            //    {
            //"Overall Length", "Datum to End", "Head Diameter", "Groove Position",
            //"Stem Dia Near Groove", "Stem Dia Near Undercut", "Groove Diameter",
            //"Straightness", "Seat Height", "Seat Runout", "Datum to Groove",
            //"Ovality SDG", "Ovality SDU", "Ovality Head", "Stem Taper",
            //"Face Runout", "End Face Runout"
            //    };

            var calculatedValues = new Dictionary<string, double>();

            // Create dummy ParameterInfo list (no sign change/compensation)
            var parameterInfos = parameterNames.Select(p => new ParameterInfo
            {
                Name = p,
                SignChange = 0,
                Compensation = 0.0
            }).ToList();

            foreach (var pInfo in parameterInfos)
            {
                try
                {
                    double val = CalculateProbeValue(pInfo, probeMeasurements, dbRefDict, mode);
                    calculatedValues[pInfo.Name] = val;
                }
                catch
                {
                    calculatedValues[pInfo.Name] = double.NaN;
                }
            }


            bool overallOk = true;
            var resultsWithStatus = new Dictionary<string, ParameterResult>();

            foreach (var paramName in parameterNames)
            {
                double val = calculatedValues.ContainsKey(paramName) ? calculatedValues[paramName] : double.NaN;

                var config = masterVals?.FirstOrDefault(m =>
                    string.Equals(m.Para_No, paramName, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(m.Parameter, paramName, StringComparison.OrdinalIgnoreCase)
                );

                double masterVal = config?.Nominal ?? 0;
                double tolPlus = config?.RTolPlus ?? 0;
                double tolMinus = config?.RTolMinus ?? 0;

                double minAllowed = masterVal - tolMinus;
                double maxAllowed = masterVal + tolPlus;

                bool isOk = !double.IsNaN(val) && val >= minAllowed && val <= maxAllowed;

                if (!isOk) overallOk = false;

                resultsWithStatus[paramName] = new ParameterResult
                {
                    Value = Math.Abs(val),
                    IsOk = isOk
                };
            }

            MasteringOK = overallOk;
            OnCalculatedValuesWithStatusReady(resultsWithStatus);

            // Debug info
            //System.Diagnostics.Debug.WriteLine("=== MASTER INSPECTION RESULTS WITH TOLERANCES ===");
            //foreach (var kvp in resultsWithStatus)
            //{
            //    System.Diagnostics.Debug.WriteLine($"{kvp.Key}: {kvp.Value.Value:F4} [{(kvp.Value.IsOk ? "OK" : "NG")}]");
            //}
        }


        private void HandleMeasurementStage(Dictionary<string, ProbeMeasurement> probeMeasurements, string partCode)
        {
            // Get reference master values
            var dbRefList = _dataStorageService.GetMasterProbeRef(partCode);
            var mode = ProcedureMode.Measurement;
            var dbRefDict = dbRefList
                .GroupBy(x => x.Name)
                .ToDictionary(g => g.Key, g => g.First().Value);

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

            var calculatedValues = new Dictionary<string, double>();

            // Calculate each probe value
            foreach (var pInfo in parameterInfos)
            {
                try
                {
                    calculatedValues[pInfo.Name] = CalculateProbeValue(pInfo, probeMeasurements, dbRefDict, mode);
                }
                catch
                {
                    calculatedValues[pInfo.Name] = double.NaN;
                }
            }

            // Determine overall OK/NG and prepare results
            bool overallOk = true;
            var resultsWithStatus = new Dictionary<string, ParameterResult>();

            foreach (var paramName in parameterNames)
            {
                double val = calculatedValues[paramName];

                var config = masterVals?.FirstOrDefault(m =>
                    string.Equals(m.Para_No, paramName, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(m.Parameter, paramName, StringComparison.OrdinalIgnoreCase)
                );

                double masterVal = config?.Nominal ?? 0;
                double tolPlus = config?.RTolPlus ?? 0;
                double tolMinus = config?.RTolMinus ?? 0;

                double minAllowed = masterVal - tolMinus;
                double maxAllowed = masterVal + tolPlus;

                bool isOk = !double.IsNaN(val) && val >= minAllowed && val <= maxAllowed;
                if (!isOk) overallOk = false;

                resultsWithStatus[paramName] = new ParameterResult
                {
                    Value = Math.Abs(val),
                    IsOk = isOk
                };
            }

            MasteringOK = overallOk;

            // Send calculated results to the Result Page
            OnCalculatedValuesWithStatusReady(resultsWithStatus);



            int bitValue = GetPlcDeviceBit("X14"); // Auto/Manual PLC bit

            var autoList = _dataStorageService.GetActiveBit();

            var shControl = autoList.FirstOrDefault(c => c.Code == "SH");
            var sroControl = autoList.FirstOrDefault(c => c.Code == "SRO");
            var stdiControl = autoList.FirstOrDefault(c => c.Code == "STDI");
            var gdControl = autoList.FirstOrDefault(c => c.Code == "GD");

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
                        SetPlcDevice("M13", 0); // Turn OFF
                    }


                    // ---- Case 1: More than 2 NGs => Direct general rejection ----
                    if (ngCount > 2)
                    {
                        SetPlcDevice("M302", 1); // General rejection
                        return; // stop checking further
                    }

                    // ---- Case 2: 1 or 2 NGs => check specific conditions ----
                    bool anyReject = false;

                    // ---- Seat Height ----
                    if (resultsWithStatus.ContainsKey("Seat Height") &&
                        !resultsWithStatus["Seat Height"].IsOk &&
                        shControl != null && shControl.Bit == 1)
                    {
                        SetPlcDevice("M305", 1); // Seat Height rejection
                        anyReject = true;
                    }

                    // ---- Seat Runout ----
                    if (resultsWithStatus.ContainsKey("Seat Runout") &&
                        !resultsWithStatus["Seat Runout"].IsOk &&
                        sroControl != null && sroControl.Bit == 1)
                    {
                        SetPlcDevice("M303", 1); // Seat Runout rejection
                        anyReject = true;
                    }

                    // ---- Stem Diameter ----
                    if (((resultsWithStatus.ContainsKey("Stem Dia Near Groove") &&
                          !resultsWithStatus["Stem Dia Near Groove"].IsOk) ||
                         (resultsWithStatus.ContainsKey("Stem Dia Near Undercut") &&
                          !resultsWithStatus["Stem Dia Near Undercut"].IsOk)) &&
                        stdiControl != null && stdiControl.Bit == 1)
                    {
                        SetPlcDevice("M304", 1); // Stem Diameter rejection
                        anyReject = true;
                    }

                    // ---- Groove ----
                    //if (((resultsWithStatus.ContainsKey("Groove Diameter") &&
                    //      !resultsWithStatus["Groove Diameter"].IsOk) ||
                    //     (resultsWithStatus.ContainsKey("Groove Position") &&
                    //      !resultsWithStatus["Groove Position"].IsOk)) &&
                    //    gdControl != null && gdControl.Bit == 1)
                    //{
                    //    SetPlcDevice("M307", 1); // Groove rejection
                    //    anyReject = true;
                    //}

                    // ---- Final decision ----
                    if (anyReject)
                    {
                        SetPlcDevice("M302", 1); // General rejection (any single failure)
                    }
                    else
                    {
                        SetPlcDevice("M306", 1); // All OK
                    }
                }
                else
                {
                    Debug.WriteLine("Manual mode active. Skipping PLC rejection logic.");
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
        private double CalculateProbeValue(
    ParameterInfo paramInfo,
    Dictionary<string, ProbeMeasurement> probeMeasurements,
    Dictionary<string, double> dbRefDict,
    ProcedureMode mode)
        {
            if (paramInfo == null || string.IsNullOrWhiteSpace(paramInfo.Name))
                return 0;

            string paramName = paramInfo.Name.ToLower();
            int signChange = paramInfo.SignChange;
            double compensation = paramInfo.Compensation;

            switch (paramName)
            {
                case "datum to end":
                    return CalculateDatumToEnd(probeMeasurements, dbRefDict, mode, signChange, compensation);
                case "overall length":
                    return CalculateOverallLength(probeMeasurements, dbRefDict, mode, signChange, compensation);
                case "head diameter":
                    return CalculateHeadDiameter(probeMeasurements, dbRefDict, mode, signChange, compensation);
                case "groove position":
                    return CalculateGroovePosition(probeMeasurements, dbRefDict, mode, signChange, compensation);
                case "stem dia near groove":
                    return CalculateStemDia1(probeMeasurements, dbRefDict, mode, signChange, compensation);
                case "stem dia near undercut":
                    return CalculateStemDia2(probeMeasurements, dbRefDict, mode, signChange, compensation);
                case "groove diameter":
                    return CalculateGrooveDia(probeMeasurements, dbRefDict, mode, signChange, compensation);
                case "straightness":
                    return CalculateReducedDia(probeMeasurements, dbRefDict, mode, signChange, compensation);
                case "end face runout":
                    return CalculateEFRO(probeMeasurements, dbRefDict, mode, signChange, compensation);
                case "seat height":
                    return CalculateSeatHeight(probeMeasurements, dbRefDict, mode, signChange, compensation);
                case "seat runout":
                    return CalculateSeatRunout(probeMeasurements, dbRefDict, mode, signChange, compensation);
                case "datum to groove":
                    return CalculateDatumToGroove(probeMeasurements, dbRefDict, mode, signChange, compensation);
                case "ovality sdg":
                    return CalculateOvalitySDG(probeMeasurements, dbRefDict, mode, signChange, compensation);
                case "ovality sdu":
                    return CalculateOvalitySDU(probeMeasurements, dbRefDict, mode, signChange, compensation); 
                case "ovality head":
                    return CalculateOvalityHead(probeMeasurements, dbRefDict, mode, signChange, compensation);
                case "stem taper":
                    return CalculateStemTaper(probeMeasurements, dbRefDict, mode, signChange, compensation);
                case "face runout":
                    return CalculateFaceRunout(probeMeasurements, dbRefDict, mode, signChange, compensation);

                default:
                    throw new Exception($"Unknown probe DB name: {paramInfo.Name}");
            }
        }


        #region Calculation Methods
        private double CalculateDatumToEnd(
            Dictionary<string, ProbeMeasurement> probeMeasurements,
            Dictionary<string, double> dbRefDict,
            ProcedureMode mode,
            int signChange,
            double compensation)
        {
            if (!probeMeasurements.TryGetValue("Datum to End", out var pm)) return double.NaN;
            if (!dbRefDict.TryGetValue("Datum to End", out var dbRefValue)) dbRefValue = 0.0;

            double current = pm.MaxValue;
            double offset = current - dbRefValue;
            double datumToEnd = Math.Abs(offset);

            if (offset > 0)
                datumToEnd += datumToEnd * 0.33;
            else if (offset < 0)
                datumToEnd += datumToEnd * 0.40;

            double result;

            if (mode == ProcedureMode.Measurement)
            {
                result = (signChange == 1)
                    ? pm.MasterValue - datumToEnd
                    : pm.MasterValue + datumToEnd;

                if (compensation != 0)
                    result += compensation;
            }
            else
            {
                result = pm.MasterValue + datumToEnd;
            }

            return Math.Round(result, 4);
        }

        private double CalculateOverallLength(
            Dictionary<string, ProbeMeasurement> probeMeasurements,
            Dictionary<string, double> dbRefDict,
            ProcedureMode mode,
            int signChange,
            double compensation)
        {
            if (!probeMeasurements.TryGetValue("Overall Length", out var pm)) return double.NaN;
            if (!dbRefDict.TryGetValue("Overall Length", out var dbRefValue)) dbRefValue = 0.0;

            double offset = pm.MaxValue - dbRefValue;
            double result;

            if (mode == ProcedureMode.Measurement)
            {
                result = (signChange == 1)
                    ? pm.MasterValue - offset
                    : pm.MasterValue + offset;

                if (compensation != 0)
                    result += compensation;
            }
            else
            {
                result = pm.MasterValue + offset;
            }

            return Math.Round(result, 3);
        }

        private double CalculateHeadDiameter(
            Dictionary<string, ProbeMeasurement> probeMeasurements,
            Dictionary<string, double> dbRefDict,
            ProcedureMode mode,
            int signChange,
            double compensation)
        {
            if (!probeMeasurements.TryGetValue("Head Diameter", out var pm)) return double.NaN;
            if (!dbRefDict.TryGetValue("Head Diameter", out var dbRefValue)) dbRefValue = 0.0;

            double offset = pm.MaxValue - dbRefValue;
            double result;

            if (mode == ProcedureMode.Measurement)
            {
                result = (signChange == 1)
                    ? pm.MasterValue - offset
                    : pm.MasterValue + offset;

                if (compensation != 0)
                    result += compensation;
            }
            else
            {
                result = pm.MasterValue + offset;
            }

            return Math.Round(result, 3);
        }

        private double CalculateGroovePosition(
            Dictionary<string, ProbeMeasurement> probeMeasurements,
            Dictionary<string, double> dbRefDict,
            ProcedureMode mode,
            int signChange,
            double compensation)
        {
            if (!probeMeasurements.TryGetValue("Groove Position", out var pm)) return double.NaN;
            if (!dbRefDict.TryGetValue("Groove Position", out var dbRefValue)) dbRefValue = 0.0;

            double offset = pm.MaxValue - dbRefValue;
            double result;

            if (mode == ProcedureMode.Measurement)
            {
                result = (signChange == 1)
                    ? pm.MasterValue - offset
                    : pm.MasterValue + offset;

                if (compensation != 0)
                    result += compensation;
            }
            else
            {
                result = pm.MasterValue + offset;
            }

            return Math.Round(result, 3);
        }

        private double CalculateStemDia1(
            Dictionary<string, ProbeMeasurement> probeMeasurements,
            Dictionary<string, double> dbRefDict,
            ProcedureMode mode,
            int signChange,
            double compensation)
        {
            if (!probeMeasurements.TryGetValue("Stem Dia Near Groove", out var pm)) return double.NaN;
            if (!dbRefDict.TryGetValue("Stem Dia Near Groove", out var dbRefValue)) dbRefValue = 0.0;

            double offset = pm.MaxValue - dbRefValue;
            double result;

            if (mode == ProcedureMode.Measurement)
            {
                result = (signChange == 1)
                    ? pm.MasterValue - offset
                    : pm.MasterValue + offset;

                if (compensation != 0)
                    result += compensation;
            }
            else
            {
                result = pm.MasterValue + offset;
            }

            return Math.Round(result, 3);
        }

        private double CalculateStemDia2(
            Dictionary<string, ProbeMeasurement> probeMeasurements,
            Dictionary<string, double> dbRefDict,
            ProcedureMode mode,
            int signChange,
            double compensation)
        {
            if (!probeMeasurements.TryGetValue("Stem Dia Near Undercut", out var pm)) return double.NaN;
            if (!dbRefDict.TryGetValue("Stem Dia Near Undercut", out var dbRefValue)) dbRefValue = 0.0;

            double offset = pm.MaxValue - dbRefValue;
            double result;

            if (mode == ProcedureMode.Measurement)
            {
                result = (signChange == 1)
                    ? pm.MasterValue - offset
                    : pm.MasterValue + offset;

                if (compensation != 0)
                    result += compensation;
            }
            else
            {
                result = pm.MasterValue + offset;
            }

            return Math.Round(result, 3);
        }

        private double CalculateReducedDia(
            Dictionary<string, ProbeMeasurement> probeMeasurements,
            Dictionary<string, double> dbRefDict,
            ProcedureMode mode,
            int signChange,
            double compensation)
        {
            if (!probeMeasurements.TryGetValue("Straightness", out var pm)) return double.NaN;
            if (!dbRefDict.TryGetValue("Straightness", out var dbRefValue)) dbRefValue = 0.0;

            double offset = pm.MaxValue - dbRefValue;
            if (offset < 0.001) offset = 0.001;

            double result;

            if (mode == ProcedureMode.Measurement)
            {
                result = (signChange == 1)
                    ? pm.MasterValue - offset
                    : pm.MasterValue + offset;

                if (compensation != 0)
                    result += compensation;
            }
            else
            {
                result = pm.MasterValue + offset;
            }

            return Math.Round(result, 3);
        }

        private double CalculateEFRO(
            Dictionary<string, ProbeMeasurement> probeMeasurements,
            Dictionary<string, double> dbRefDict,
            ProcedureMode mode,
            int signChange,
            double compensation)
        {
            if (!probeMeasurements.TryGetValue("End Face Runout", out var pm)) return double.NaN;
            if (!dbRefDict.TryGetValue("End Face Runout", out var dbRefValue)) dbRefValue = 0.0;

            double offset = pm.MaxValue - dbRefValue;
            double result;

            if (mode == ProcedureMode.Measurement)
            {
                result = (signChange == 1)
                    ? pm.MasterValue - offset
                    : pm.MasterValue + offset;

                if (compensation != 0)
                    result += compensation;
            }
            else
            {
                result = pm.MasterValue + offset;
            }

            return Math.Round(result, 3);
        }

        private double CalculateGrooveDia(
            Dictionary<string, ProbeMeasurement> probeMeasurements,
            Dictionary<string, double> dbRefDict,
            ProcedureMode mode,
            int signChange,
            double compensation)
        {
            if (!probeMeasurements.TryGetValue("Groove Diameter", out var pm)) return double.NaN;
            if (!dbRefDict.TryGetValue("Groove Diameter", out var dbRefValue)) dbRefValue = 0.0;

            double offset = pm.MaxValue - dbRefValue;
            double result;

            if (mode == ProcedureMode.Measurement)
            {
                result = (signChange == 1)
                    ? pm.MasterValue - offset
                    : pm.MasterValue + offset;

                if (compensation != 0)
                    result += compensation;
            }
            else
            {
                result = pm.MasterValue + offset;
            }

            return Math.Round(result, 3);
        }

        private double CalculateSeatHeight(
            Dictionary<string, ProbeMeasurement> probeMeasurements,
            Dictionary<string, double> dbRefDict,
            ProcedureMode mode,
            int signChange,
            double compensation)
        {
            double ol = CalculateOverallLength(probeMeasurements, dbRefDict, mode, signChange, compensation);
            double de = CalculateDatumToEnd(probeMeasurements, dbRefDict, mode, signChange, compensation);

            if (double.IsNaN(ol) || double.IsNaN(de)) return double.NaN;
            return Math.Round(ol - de, 3);
        }

        private double CalculateSeatRunout(
            Dictionary<string, ProbeMeasurement> probeMeasurements,
            Dictionary<string, double> dbRefDict,
            ProcedureMode mode,
            int signChange,
            double compensation)
        {
            if (!probeMeasurements.TryGetValue("Datum to End", out var pm)) return double.NaN;

            double seatRunout = pm.MaxValue - pm.MinValue;

            if (mode == ProcedureMode.Measurement && compensation != 0)
                seatRunout += compensation;

            return Math.Round(seatRunout, 3);
        }

        private double CalculateDatumToGroove(
            Dictionary<string, ProbeMeasurement> probeMeasurements,
            Dictionary<string, double> dbRefDict,
            ProcedureMode mode,
            int signChange,
            double compensation)
        {
            double de = CalculateDatumToEnd(probeMeasurements, dbRefDict, mode, signChange, compensation);
            double gp = CalculateGroovePosition(probeMeasurements, dbRefDict, mode, signChange, compensation);

            if (double.IsNaN(de) || double.IsNaN(gp)) return double.NaN;
            return Math.Round(de - gp, 3);
        }

        private double CalculateOvalitySDG(
            Dictionary<string, ProbeMeasurement> probeMeasurements,
            Dictionary<string, double> dbRefDict,
            ProcedureMode mode,
            int signChange,
            double compensation)
        {
            if (!probeMeasurements.TryGetValue("Stem Dia Near Groove", out var pm)) return double.NaN;

            double result = pm.MaxValue - pm.MinValue;

            if (mode == ProcedureMode.Measurement && compensation != 0)
                result += compensation;

            return Math.Round(result, 3);
        }

        private double CalculateOvalitySDU(
            Dictionary<string, ProbeMeasurement> probeMeasurements,
            Dictionary<string, double> dbRefDict,
            ProcedureMode mode,
            int signChange,
            double compensation)
        {
            if (!probeMeasurements.TryGetValue("Stem Dia Near Undercut", out var pm)) return double.NaN;

            double result = pm.MaxValue - pm.MinValue;
            if (result < 0.001) result = 0.001;

            if (mode == ProcedureMode.Measurement && compensation != 0)
                result += compensation;

            return Math.Round(result, 3);
        }

        private double CalculateOvalityHead(
            Dictionary<string, ProbeMeasurement> probeMeasurements,
            Dictionary<string, double> dbRefDict,
            ProcedureMode mode,
            int signChange,
            double compensation)
        {
            if (!probeMeasurements.TryGetValue("Head Diameter", out var pm)) return double.NaN;

            double ovality = pm.MaxValue - pm.MinValue;

            if (mode == ProcedureMode.Measurement && compensation != 0)
                ovality += compensation;

            return Math.Round(ovality, 3);
        }

        private double CalculateStemTaper(
            Dictionary<string, ProbeMeasurement> probeMeasurements,
            Dictionary<string, double> dbRefDict,
            ProcedureMode mode,
            int signChange,
            double compensation)
        {
            double stemDia1 = CalculateStemDia1(probeMeasurements, dbRefDict, mode, signChange, compensation);
            double stemDia2 = CalculateStemDia2(probeMeasurements, dbRefDict, mode, signChange, compensation);

            if (double.IsNaN(stemDia1) || double.IsNaN(stemDia2))
                return double.NaN;

            double stemTaper = stemDia1 - stemDia2;
            return Math.Round(stemTaper, 3);
        }

        private double CalculateFaceRunout(
            Dictionary<string, ProbeMeasurement> probeMeasurements,
            Dictionary<string, double> dbRefDict,
            ProcedureMode mode,
            int signChange,
            double compensation)
        {
            if (!probeMeasurements.TryGetValue("Overall Length", out var pm)) return double.NaN;

            double faceRunout = pm.MaxValue - pm.MinValue;

            if (mode == ProcedureMode.Measurement && compensation != 0)
                faceRunout += compensation;

            return Math.Round(faceRunout, 3);
        }


        #endregion




        private void LoadProbeConfigurations(string partCode)
        {
            _probeMeasurements.Clear();
            _orderedProbeMeasurements.Clear();

            var probeInstalls = _dataStorageService.GetProbeInstallByPartNumber(partCode);
            var masterVals = _dataStorageService.GetPartConfig(partCode); // <-- Updated here

            foreach (var probe in probeInstalls)
            {
                var config = masterVals?
                    .FirstOrDefault(m =>
                        string.Equals(m.Para_No, probe.ProbeId, StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(m.Parameter, probe.Name, StringComparison.OrdinalIgnoreCase)
                    );

                double masterVal = config?.Nominal ?? 0;
                double tolPlus = config?.RTolPlus ?? 0;
                double tolMinus = config?.RTolMinus ?? 0;
                int Sign= config?.Sign_Change ?? 0;
                double Comp=config?.Compensation ?? 0;

                var pm = new ProbeMeasurement
                {
                    ProbeId = probe.ProbeId,
                    Name = probe.Name,
                    MasterValue = masterVal,
                    TolerancePlus = tolPlus,
                    ToleranceMinus = tolMinus,
                    SignChange= Sign,
                    Compensation= Comp,
                   
                };

                _probeMeasurements[probe.ProbeId] = pm;
                _orderedProbeMeasurements.Add(pm); // Keep ordered list in database order
            }

            System.Diagnostics.Debug.WriteLine($"[DEBUG] Loaded {_probeMeasurements.Count} probe configs for part {partCode}");
            foreach (var pm in _orderedProbeMeasurements)
            {
                System.Diagnostics.Debug.WriteLine($"  Probe {pm.ProbeId} ({pm.Name}): Master={pm.MasterValue}, Tol±={pm.TolerancePlus}/{pm.ToleranceMinus}");
            }
        }


        private void LoadProbeConfigurationsforMasterInspection(string partCode)
        {
            _probeMeasurements.Clear();
            _orderedProbeMeasurements.Clear();

            var probeInstalls = _dataStorageService.GetProbeInstallByPartNumber(partCode);
            var masterVals = _dataStorageService.GetMasterReadingByPart(partCode); // <-- Updated here

            foreach (var probe in probeInstalls)
            {
                var config = masterVals?
                    .FirstOrDefault(m =>
                        string.Equals(m.Para_No, probe.ProbeId, StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(m.Parameter, probe.Name, StringComparison.OrdinalIgnoreCase)
                    );

                double masterVal = config?.Nominal ?? 0;
                double tolPlus = config?.RTolPlus ?? 0;
                double tolMinus = config?.RTolMinus ?? 0;
                

                var pm = new ProbeMeasurement
                {
                    ProbeId = probe.ProbeId,
                    Name = probe.Name,
                    MasterValue = masterVal,
                    TolerancePlus = tolPlus,
                    ToleranceMinus = tolMinus,
                };

                _probeMeasurements[probe.ProbeId] = pm;
                _orderedProbeMeasurements.Add(pm); // Keep ordered list in database order
            }

            System.Diagnostics.Debug.WriteLine($"[DEBUG] Loaded {_probeMeasurements.Count} probe configs for part {partCode}");
            foreach (var pm in _orderedProbeMeasurements)
            {
                System.Diagnostics.Debug.WriteLine($"  Probe {pm.ProbeId} ({pm.Name}): Master={pm.MasterValue}, Tol±={pm.TolerancePlus}/{pm.ToleranceMinus}");
            }
        }



        public void Dispose()
        {
            Cleanup();
            _dataStorageService?.Dispose();
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
        //        
        //public async Task WaitForValidProbeReadingAsync(string targetProbeId, bool suppressDetectedMessage = false)
        //{
        //    await NotifyOnUIAsync("Initializing probe readings...");

        //    _plcProbeService.StartLiveReading(100);

        //    bool partDetected = false;
        //    bool messageFired = false;

        //    if (!suppressDetectedMessage)
        //        if (GetPlcDeviceBit("X14") == 1)
        //            await NotifyOnUIAsync("Waiting the robo to load part");
        //        else
        //            await NotifyOnUIAsync("Load the part");
        //    while (!partDetected)
        //    {
        //        await Task.Delay(100);

        //        var probeVal = _collectedReadings
        //            .Where(r => r.ProbeId == targetProbeId)
        //            .Select(r => r.Value)
        //            .LastOrDefault();

        //        bool partPresent = Math.Abs(probeVal) > 0.100;

        //        if (partPresent && !messageFired)
        //        {
        //            messageFired = true;
        //            partDetected = true;
        //            if (!suppressDetectedMessage)
        //                await NotifyOnUIAsync("Part detected. Proceeding...");
        //        }
        //        else if (!partPresent && messageFired)
        //        {
        //            messageFired = false;
        //            if (!suppressDetectedMessage)
        //                await NotifyOnUIAsync("Part removed. Waiting for new part...");
        //        }
        //        else if (!partPresent && !messageFired)
        //        {
        //            if (!suppressDetectedMessage)
        //                if (GetPlcDeviceBit("X14") == 1)
        //                    await NotifyOnUIAsync("Waiting the robo to load part");
        //                else
        //                    await NotifyOnUIAsync("Load the part");
        //        }
        //    }

        //    _plcProbeService.StopLiveReading();

        //    if (!suppressDetectedMessage)
        //        await NotifyOnUIAsync("Values updated...");
        //}


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
