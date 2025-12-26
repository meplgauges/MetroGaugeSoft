using ActUtlType64Lib;
using System.Collections.Concurrent;
using System.Windows;

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
        private string _currentPartCode = "3366566";

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

            LoadProbeConfigurationsforMasterInspection(_currentPartCode);

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
                if (!connected) return false;
            }

            // PLC connection (unchanged)
            int openResult = plc.Open();
            if (openResult != 0)
            {
                MessageBox.Show($"Failed to open PLC connection. Error code: {openResult}",
                    "PLC Init Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }

            return true;
        }


        //private void ProbeValueUpdatedHandler(object? sender, ProbeReadingEventArgs e)
        //{
        //    _collectedReadings.Enqueue((e.ModuleId, e.Value));
        //}

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
                if (mode == ProcedureMode.MasterInspection)
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


        public async Task WaitForValidProbeReadingAsync(string targetParameterName, bool suppressDetectedMessage = false)
        {
            await NotifyOnUIAsync("Initializing probe readings...");

            _plcProbeService.StartLiveReading(100);  // ✅ Start serial reading

            bool partDetected = false;
            bool messageFired = false;

            if (!suppressDetectedMessage)
            {
                if (GetPlcDeviceBit("X14") == 1)
                    await NotifyOnUIAsync("Waiting the robo to load part");
                else
                    await NotifyOnUIAsync("Load the part");
            }

            while (!partDetected)
            {
                await Task.Delay(100);

                // ✅ CHANGE: Direct dictionary access
                double? probeVal = _plcProbeService.GetProbeValue(targetParameterName);
                bool partPresent = Math.Abs(probeVal ?? 0) > 0.100;

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
            try
            {
                NotifyStatus("Collecting probe readings...");

                // Reset previous data
                foreach (var pm in sortedProbeMeasurements)
                    pm.Readings.Clear();

                while (_collectedReadings.TryDequeue(out _)) { }

                _plcProbeService.StartLiveReading(15);

                int completed = 0;

                while (!Abort)
                {
                    // DEQUEUE from probe service
                    while (_plcProbeService.CollectedReadings.TryDequeue(out var reading))
                    {
                        if (_probeMeasurements.TryGetValue(reading.ParameterName, out var pm))
                        {
                            if (pm.Readings.Count < ArraySize)
                            {
                                pm.Readings.Add(reading.Value);

                                if (pm.Readings.Count == ArraySize)
                                    completed++;
                            }
                        }
                    }

                    if (completed >= sortedProbeMeasurements.Count)
                        break;

                    await Task.Delay(10);
                }

                _plcProbeService.StopLiveReading();
                NotifyStatus("Live reading stopped. Processing results...");

                // POST-PROCESSING
                foreach (var pm in sortedProbeMeasurements)
                {
                    if (pm.Readings.Count > 0)
                    {
                        var clean = pm.Readings.Where(x => !double.IsNaN(x)).ToList();
                        pm.MinValue = clean.Min();
                        pm.MaxValue = clean.Max();
                    }
                    else
                    {
                        pm.MinValue = pm.MaxValue = 0;
                    }
                }

                var probeMeasurementByNameMeasurement = sortedProbeMeasurements
                   .Where(pm => !string.IsNullOrEmpty(pm.Name))
                   .ToDictionary(pm => pm.Name!);
                NotifyStatus("Probe data collection completed successfully.");
                // await HandleProcedurePostProcessingAsync(ProcedureMode.Mastering, _orderedProbeMeasurements, true);
                HandleMeasurementStage(probeMeasurementByNameMeasurement, _currentPartCode);


            }
            catch (Exception ex)
            {
                NotifyStatus($"Error during probe reading: {ex.Message}");
                _plcProbeService.StopLiveReading();
            }
        }





        private async Task HandleProcedurePostProcessingAsync(ProcedureMode mode, List<ProbeMeasurement> probeMeasurements, bool isFirstMeasurementCycle)
        {
            switch (mode)
            {
                case ProcedureMode.Mastering:
                    await NotifyOnUIAsync("Mastering Completed. Press Enter to Inspect the Master");

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
            var probeMeasurementByName =
                                         BuildEffectiveProbesByName(probeMeasurements.Values.ToList());
            var dbRefList = _dataStorageService.GetMasterProbeRef(_currentPartCode);

            var mode = ProcedureMode.MasterInspection;

            // ✅ FIXED: Prevent "duplicate key" crash
            var dbRefDict = dbRefList
                 .GroupBy(x => x.Name)
                 .ToDictionary(
                     g => g.Key,
                     g => (Min: g.First().Min, Max: g.First().Max));

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
                    double val = CalculateProbeValue(pInfo, probeMeasurementByName, dbRefDict, mode);
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

                double FinalResult = val + masterVal;

                bool isOk = !double.IsNaN(FinalResult) && FinalResult >= minAllowed && FinalResult <= maxAllowed;
                if (!isOk) overallOk = false;

                resultsWithStatus[paramName] = new ParameterResult
                {
                    Value = Math.Abs(FinalResult),
                    IsOk = isOk
                };
            }

            MasteringOK = overallOk;

            // Send calculated results to the Result Page
            OnCalculatedValuesWithStatusReady(resultsWithStatus);




            //int bitValue = GetPlcDeviceBit("X14"); // Auto/Manual PLC bit

            //var autoList = _dataStorageService.GetActiveBit();

            //var shControl = autoList.FirstOrDefault(c => c.Code == "SH");
            //var sroControl = autoList.FirstOrDefault(c => c.Code == "SRO");
            //var stdiControl = autoList.FirstOrDefault(c => c.Code == "STDI");
            //var gdControl = autoList.FirstOrDefault(c => c.Code == "GD");

            //try
            //{
            //    if (bitValue == 1) // Auto mode only
            //    {
            //        int ngCount = resultsWithStatus.Count(r => !r.Value.IsOk);

            //        // If any NG exists, pulse signal M13
            //        if (ngCount > 0)
            //        {
            //            SetPlcDevice("M13", 1); // Turn ON
            //            Thread.Sleep(10);        // 5 ms pulse (adjust as needed)
            //            SetPlcDevice("M13", 0); // Turn OFF
            //        }


            //        // ---- Case 1: More than 2 NGs => Direct general rejection ----
            //        if (ngCount > 2)
            //        {
            //            SetPlcDevice("M302", 1); // General rejection
            //            return; // stop checking further
            //        }

            //        // ---- Case 2: 1 or 2 NGs => check specific conditions ----
            //        bool anyReject = false;

            //        // ---- Seat Height ----
            //        if (resultsWithStatus.ContainsKey("Seat Height") &&
            //            !resultsWithStatus["Seat Height"].IsOk &&
            //            shControl != null && shControl.Bit == 1)
            //        {
            //            SetPlcDevice("M305", 1); // Seat Height rejection
            //            anyReject = true;
            //        }

            //        // ---- Seat Runout ----
            //        if (resultsWithStatus.ContainsKey("Seat Runout") &&
            //            !resultsWithStatus["Seat Runout"].IsOk &&
            //            sroControl != null && sroControl.Bit == 1)
            //        {
            //            SetPlcDevice("M303", 1); // Seat Runout rejection
            //            anyReject = true;
            //        }

            //        // ---- Stem Diameter ----
            //        if (((resultsWithStatus.ContainsKey("Stem Dia Near Groove") &&
            //              !resultsWithStatus["Stem Dia Near Groove"].IsOk) ||
            //             (resultsWithStatus.ContainsKey("Stem Dia Near Undercut") &&
            //              !resultsWithStatus["Stem Dia Near Undercut"].IsOk)) &&
            //            stdiControl != null && stdiControl.Bit == 1)
            //        {
            //            SetPlcDevice("M304", 1); // Stem Diameter rejection
            //            anyReject = true;
            //        }

            //        // ---- Groove ----
            //        //if (((resultsWithStatus.ContainsKey("Groove Diameter") &&
            //        //      !resultsWithStatus["Groove Diameter"].IsOk) ||
            //        //     (resultsWithStatus.ContainsKey("Groove Position") &&
            //        //      !resultsWithStatus["Groove Position"].IsOk)) &&
            //        //    gdControl != null && gdControl.Bit == 1)
            //        //{
            //        //    SetPlcDevice("M307", 1); // Groove rejection
            //        //    anyReject = true;
            //        //}

            //        // ---- Final decision ----
            //        if (anyReject)
            //        {
            //            SetPlcDevice("M302", 1); // General rejection (any single failure)
            //        }
            //        else
            //        {
            //            SetPlcDevice("M306", 1); // All OK
            //        }
            //    }
            //    else
            //    {
            //        Debug.WriteLine("Manual mode active. Skipping PLC rejection logic.");
            //    }
            //}
            //catch (Exception ex)
            //{
            //    Debug.WriteLine($"Error setting PLC bits for rejection: {ex.Message}");
            //}

        }

        protected virtual void OnCalculatedValuesReady(Dictionary<string, double> calculatedValues)
        {
            CalculatedValuesReady?.Invoke(this, calculatedValues);
        }



        // Calculation dispatcher adapted to accept both probeMeasurements and dbRefDict
        private double CalculateProbeValue(
     ParameterInfo paramInfo,
     Dictionary<string, ProbeMeasurement> probeMeasurements,
     Dictionary<string, (double Min, double Max)> dbRefDict,
     ProcedureMode mode)
        {
            if (paramInfo == null || string.IsNullOrWhiteSpace(paramInfo.Name))
                return 0;

            string paramName = paramInfo.Name.ToLower();
            int signChange = paramInfo.SignChange;
            double compensation = paramInfo.Compensation;

            switch (paramName)
            {
                case "step od1": return CalculateStepOD1(probeMeasurements, dbRefDict, mode, signChange, compensation);
                case "step od2": return CalculateStepOD2(probeMeasurements, dbRefDict, mode, signChange, compensation);
                case "od-1": return CalculateOD1(probeMeasurements, dbRefDict, mode, signChange, compensation);
                case "od-2": return CalculateOD2(probeMeasurements, dbRefDict, mode, signChange, compensation);
                case "od-3": return CalculateOD3(probeMeasurements, dbRefDict, mode, signChange, compensation);
                case "id-1": return CalculateID1(probeMeasurements, dbRefDict, mode, signChange, compensation);
                case "id-2": return CalculateID2(probeMeasurements, dbRefDict, mode, signChange, compensation);
                case "ol": return CalculateOverallLength(probeMeasurements, dbRefDict, mode, signChange, compensation);
                case "step runout-1": return CalculateStepRunout1(probeMeasurements, dbRefDict, mode, signChange, compensation);
                case "step runout-2": return CalculateStepRunout2(probeMeasurements, dbRefDict, mode, signChange, compensation);
                case "rn-1": return CalculateRN1(probeMeasurements, dbRefDict, mode, signChange, compensation);
                case "rn-2": return CalculateRN2(probeMeasurements, dbRefDict, mode, signChange, compensation);
                case "rn-3": return CalculateRN3(probeMeasurements, dbRefDict, mode, signChange, compensation);
                case "rn-4": return CalculateRN4(probeMeasurements, dbRefDict, mode, signChange, compensation);
                case "rn-5": return CalculateRN5(probeMeasurements, dbRefDict, mode, signChange, compensation);
                default: throw new Exception($"Unknown probe DB name: {paramInfo.Name}");
            }
        }

        #region 🔥 MAIN PARAMETERS (Summing Live & Master from 2 Probes)

        private double CalculateStepOD1(
      Dictionary<string, ProbeMeasurement> probeMeasurements,
      Dictionary<string, (double Min, double Max)> dbRefDict,
      ProcedureMode mode,
      int signChange,
      double compensation)
        {
            // Probe 1 + Probe 2
            var p1 = probeMeasurements.Values.FirstOrDefault(p => p.ProbeId == "Probe 1" && p.Readings.Any());
            var p2 = probeMeasurements.Values.FirstOrDefault(p => p.ProbeId == "Probe 2" && p.Readings.Any());

            // Use stored reference Max values (or Min, depending on your spec)
            double PM1 = dbRefDict.TryGetValue("STEP OD1", out var ref1) ? ref1.Max : 0.0;
            double PM2 = dbRefDict.TryGetValue("STEP RUNOUT-1", out var ref2) ? ref2.Max : 0.0;

            if (p1 == null && p2 == null) return double.NaN;

            double liveValue = 0, masterValue = 0;
            if (p1 != null) liveValue += p1.Readings.Max();
            if (p2 != null) liveValue += p2.Readings.Max();

            double dbRefValue = PM1 + PM2;
            double offset = liveValue - dbRefValue;
            double result;

            if (mode == ProcedureMode.Measurement)
            {
                result = (signChange == 1) ? masterValue - offset : masterValue + offset;
                if (compensation != 0) result += compensation;
            }
            else
            {
                result = masterValue + offset;
            }

            return Math.Round(result, 3);
        }

        private double CalculateStepOD2(Dictionary<string, ProbeMeasurement> probeMeasurements, Dictionary<string, (double Min, double Max)> dbRefDict, ProcedureMode mode, int signChange, double compensation)
        {
            // Probe 9 + Probe 10

            var p1 = probeMeasurements.Values.FirstOrDefault(p => p.ProbeId == "Probe 9" && p.Readings.Any());
            var p2 = probeMeasurements.Values.FirstOrDefault(p => p.ProbeId == "Probe 10" && p.Readings.Any());

            double PM1 = dbRefDict.TryGetValue("STEP OD2", out var refValue) ? refValue.Max : 0.0;
            double PM2 = dbRefDict.TryGetValue("STEP RUNOUT-2", out var refValue1) ? refValue.Max : 0.0;

            double dbRefValue = PM1 + PM2;
            if (p1 == null && p2 == null) return double.NaN;

            double liveValue = 0, masterValue = 0;
            if (p1 != null)
            {
                liveValue += p1.Readings.Max();
            }
            if (p2 != null)
            {
                liveValue += p2.Readings.Max();
            }

            double offset = liveValue - dbRefValue;
            double result;

            if (mode == ProcedureMode.Measurement)
            {
                result = (signChange == 1) ? masterValue - offset : masterValue + offset;
                if (compensation != 0) result += compensation;
            }
            else
            {
                result = masterValue + offset;
            }
            return Math.Abs(Math.Round(result, 3));
        }

        private double CalculateOD1(Dictionary<string, ProbeMeasurement> probeMeasurements, Dictionary<string, (double Min, double Max)> dbRefDict, ProcedureMode mode, int signChange, double compensation)
        {
            // Probe 3 + Probe 4
            var p1 = probeMeasurements.Values.FirstOrDefault(p => p.ProbeId == "Probe 3" && p.Readings.Any());
            var p2 = probeMeasurements.Values.FirstOrDefault(p => p.ProbeId == "Probe 4" && p.Readings.Any());

            double PM1 = dbRefDict.TryGetValue("OD-1", out var refValue) ? refValue.Max : 0.0;
            double PM2 = dbRefDict.TryGetValue("RN-1", out var refValue1) ? refValue.Max : 0.0;

            if (p1 == null && p2 == null) return double.NaN;

            double liveValue = 0, masterValue = 0;
            if (p1 != null)
            {
                liveValue += p1.Readings.Max();
            }
            if (p2 != null)
            {
                liveValue += p2.Readings.Max();
            }

            double dbRefValue = PM1 + PM2;
            double offset = liveValue - dbRefValue;
            double result;

            if (mode == ProcedureMode.Measurement)
            {
                result = (signChange == 1) ? masterValue - offset : masterValue + offset;
                if (compensation != 0) result += compensation;
            }
            else
            {
                result = masterValue + offset;
            }
            return Math.Abs(Math.Round(result, 3));
        }

        private double CalculateOD2(Dictionary<string, ProbeMeasurement> probeMeasurements, Dictionary<string, (double Min, double Max)> dbRefDict, ProcedureMode mode, int signChange, double compensation)
        {
            // Probe 5 + Probe 6
            var p1 = probeMeasurements.Values.FirstOrDefault(p => p.ProbeId == "Probe 5" && p.Readings.Any());
            var p2 = probeMeasurements.Values.FirstOrDefault(p => p.ProbeId == "Probe 6" && p.Readings.Any());
            double MP1 = dbRefDict.TryGetValue("OD-2", out var refValue) ? refValue.Max : 0.0;
            double MP2 = dbRefDict.TryGetValue("RN-2", out var refValue1) ? refValue.Max : 0.0;

            if (p1 == null && p2 == null) return double.NaN;

            double liveValue = 0, masterValue = 0;
            if (p1 != null) { liveValue += p1.Readings.Max(); }
            if (p2 != null) { liveValue += p2.Readings.Max(); }

            double dbRefValue = MP1 + MP2;
            double offset = liveValue - dbRefValue;
            double result;

            if (mode == ProcedureMode.Measurement)
            {
                result = (signChange == 1) ? masterValue - offset : masterValue + offset;
                if (compensation != 0) result += compensation;
            }
            else
            {
                result = masterValue + offset;
            }
            return Math.Abs(Math.Round(result, 3));
        }

        private double CalculateOD3(Dictionary<string, ProbeMeasurement> probeMeasurements, Dictionary<string, (double Min, double Max)> dbRefDict, ProcedureMode mode, int signChange, double compensation)
        {
            // Probe 7 + Probe 8
            var p1 = probeMeasurements.Values.FirstOrDefault(p => p.ProbeId == "Probe 7" && p.Readings.Any());
            var p2 = probeMeasurements.Values.FirstOrDefault(p => p.ProbeId == "Probe 8" && p.Readings.Any());

            double MP1 = dbRefDict.TryGetValue("OD-3", out var refValue) ? refValue.Max : 0.0;
            double MP2 = dbRefDict.TryGetValue("RN-3", out var refValue1) ? refValue.Max : 0.0;

            if (p1 == null && p2 == null) return double.NaN;

            double liveValue = 0, masterValue = 0;
            if (p1 != null) { liveValue += p1.Readings.Max(); }
            if (p2 != null) { liveValue += p2.Readings.Max(); }

            double dbRefValue = MP1 + MP2;
            double offset = liveValue - dbRefValue;
            double result;

            if (mode == ProcedureMode.Measurement)
            {
                result = (signChange == 1) ? masterValue - offset : masterValue + offset;
                if (compensation != 0) result += compensation;
            }
            else
            {
                result = masterValue + offset;
            }
            return Math.Abs(Math.Round(result, 3));
        }

        private double CalculateID1(Dictionary<string, ProbeMeasurement> probeMeasurements, Dictionary<string, (double Min, double Max)> dbRefDict, ProcedureMode mode, int signChange, double compensation)
        {
            // Probe 11 + Probe 12
            var p1 = probeMeasurements.Values.FirstOrDefault(p => p.ProbeId == "Probe 11" && p.Readings.Any());
            var p2 = probeMeasurements.Values.FirstOrDefault(p => p.ProbeId == "Probe 12" && p.Readings.Any());

            double MP1 = dbRefDict.TryGetValue("ID-1", out var refValue) ? refValue.Max : 0.0;
            double MP2 = dbRefDict.TryGetValue("RN-4", out var refValue1) ? refValue.Max : 0.0;
            if (p1 == null && p2 == null) return double.NaN;

            double liveValue = 0, masterValue = 0;
            if (p1 != null) { liveValue += p1.Readings.Max(); masterValue += p1.MasterValue; }
            if (p2 != null) { liveValue += p2.Readings.Max(); masterValue += p2.MasterValue; }

            double dbRefValue = MP1 + MP2;
            double offset = liveValue - dbRefValue;
            double result;

            if (mode == ProcedureMode.Measurement)
            {
                result = (signChange == 1) ? masterValue - offset : masterValue + offset;
                if (compensation != 0) result += compensation;
            }
            else
            {
                result = masterValue + offset;
            }
            return Math.Abs(Math.Round(result, 3));
        }

        private double CalculateID2(Dictionary<string, ProbeMeasurement> probeMeasurements, Dictionary<string, (double Min, double Max)> dbRefDict, ProcedureMode mode, int signChange, double compensation)
        {
            // Probe 13 + Probe 14
            var p1 = probeMeasurements.Values.FirstOrDefault(p => p.ProbeId == "Probe 13" && p.Readings.Any());
            var p2 = probeMeasurements.Values.FirstOrDefault(p => p.ProbeId == "Probe 14" && p.Readings.Any());

            double MP1 = dbRefDict.TryGetValue("OD-2", out var refValue) ? refValue.Max : 0.0;
            double MP2 = dbRefDict.TryGetValue("RN-5", out var refValue1) ? refValue.Max : 0.0;
            if (p1 == null && p2 == null) return double.NaN;

            double liveValue = 0, masterValue = 0;
            if (p1 != null) { liveValue += p1.Readings.Max(); masterValue += p1.MasterValue; }
            if (p2 != null) { liveValue += p2.Readings.Max(); masterValue += p2.MasterValue; }

            double dbRefValue = MP1 + MP2;
            double offset = liveValue - dbRefValue;
            double result;

            if (mode == ProcedureMode.Measurement)
            {
                result = (signChange == 1) ? masterValue - offset : masterValue + offset;
                if (compensation != 0) result += compensation;
            }
            else
            {
                result = masterValue + offset;
            }
            return Math.Abs(Math.Round(result, 3));
        }

        private double CalculateOverallLength(Dictionary<string, ProbeMeasurement> probeMeasurements, Dictionary<string, (double Min, double Max)> dbRefDict, ProcedureMode mode, int signChange, double compensation)
        {
            // Probe 15 (Single or + Probe 16 if needed, assumed single here as per list)
            var p1 = probeMeasurements.Values.FirstOrDefault(p => p.ProbeId == "Probe 15" && p.Readings.Any());
            // Assuming Probe 16 logic if applicable, otherwise just 15

            if (p1 == null) return double.NaN;

            double liveValue = 0, masterValue = 0;
            if (p1 != null) { liveValue += p1.Readings.Max(); masterValue += p1.MasterValue; }

            double dbRefValue = dbRefDict.TryGetValue("OL", out var refValue) ? refValue.Max : 0.0;
            double offset = liveValue - dbRefValue;
            double result;

            if (mode == ProcedureMode.Measurement)
            {
                result = (signChange == 1) ? masterValue - offset : masterValue + offset;
                if (compensation != 0) result += compensation;
            }
            else
            {
                result = masterValue + offset;
            }
            return Math.Abs(Math.Round(result, 3));
        }

        #endregion

        #region 🔥 RUNOUT PARAMETERS (Max - Min, uses individual probes)

        private double CalculateStepRunout1(Dictionary<string, ProbeMeasurement> probeMeasurements, Dictionary<string, (double Min, double Max)> dbRefDict, ProcedureMode mode, int signChange, double compensation)
        {
            // Probe 2
            var pm = probeMeasurements.Values.FirstOrDefault(p => p.ProbeId == "Probe 2");
            double PM2 = dbRefDict.TryGetValue("STEP RUNOUT-1", out var refValue1) ? refValue1.Min : 0.0;


            if (pm?.Readings.Any() != true) return double.NaN;

            double MaxReading = pm.Readings.Max();

            double runout = MaxReading - PM2;

            if (runout <= 0)
            {
                runout = 0.001;
            }
            if (mode == ProcedureMode.Measurement && compensation != 0) runout += compensation;
            return Math.Abs(Math.Round(runout, 3));
        }


        private double CalculateStepRunout2(Dictionary<string, ProbeMeasurement> probeMeasurements, Dictionary<string, (double Min, double Max)> dbRefDict, ProcedureMode mode, int signChange, double compensation)
        {
            // Probe 2
            var pm = probeMeasurements.Values.FirstOrDefault(p => p.ProbeId == "Probe 10");
            double PM2 = dbRefDict.TryGetValue("STEP RUNOUT-2", out var refValue1) ? refValue1.Min : 0.0;


            if (pm?.Readings.Any() != true) return double.NaN;

            double MaxReading = pm.Readings.Min();

            double runout = MaxReading - PM2;

            if (runout < 0)
            {
                runout = 0.001;
            }
            if (mode == ProcedureMode.Measurement && compensation != 0) runout += compensation;
            return Math.Abs(Math.Round(runout, 3));

        }

        private double CalculateRN1(Dictionary<string, ProbeMeasurement> probeMeasurements, Dictionary<string, (double Min, double Max)> dbRefDict, ProcedureMode mode, int signChange, double compensation)
        {
            // Probe 4
            var pm = probeMeasurements.Values.FirstOrDefault(p => p.ProbeId == "Probe 4");
            double PM2 = dbRefDict.TryGetValue("RN-1", out var refValue1) ? refValue1.Min : 0.0;


            if (pm?.Readings.Any() != true) return double.NaN;

            double MaxReading = pm.Readings.Min();

            double runout = MaxReading - PM2;

            if (runout < 0)
            {
                runout = 0.001;
            }
            if (mode == ProcedureMode.Measurement && compensation != 0) runout += compensation;
            return Math.Abs(Math.Round(runout, 3));

        }


        private double CalculateRN2(Dictionary<string, ProbeMeasurement> probeMeasurements, Dictionary<string, (double Min, double Max)> dbRefDict, ProcedureMode mode, int signChange, double compensation)
        {
            // Probe 6
            var pm = probeMeasurements.Values.FirstOrDefault(p => p.ProbeId == "Probe 6");
            double PM2 = dbRefDict.TryGetValue("RN-2", out var refValue1) ? refValue1.Min : 0.0;


            if (pm?.Readings.Any() != true) return double.NaN;

            double MaxReading = pm.Readings.Min();

            double runout = MaxReading - PM2;

            if (runout < 0)
            {
                runout = 0.001;
            }
            if (mode == ProcedureMode.Measurement && compensation != 0) runout += compensation;
            return Math.Abs(Math.Round(runout, 3));

        }

        private double CalculateRN3(Dictionary<string, ProbeMeasurement> probeMeasurements, Dictionary<string, (double Min, double Max)> dbRefDict, ProcedureMode mode, int signChange, double compensation)
        {
            // Probe 8
            var pm = probeMeasurements.Values.FirstOrDefault(p => p.ProbeId == "Probe 8");
            double PM2 = dbRefDict.TryGetValue("RN-3", out var refValue1) ? refValue1.Min : 0.0;


            if (pm?.Readings.Any() != true) return double.NaN;

            double MaxReading = pm.Readings.Min();

            double runout = MaxReading - PM2;

            if (runout < 0)
            {
                runout = 0.001;
            }
            if (mode == ProcedureMode.Measurement && compensation != 0) runout += compensation;
            return Math.Abs(Math.Round(runout, 3));
        }


        private double CalculateRN4(Dictionary<string, ProbeMeasurement> probeMeasurements, Dictionary<string, (double Min, double Max)> dbRefDict, ProcedureMode mode, int signChange, double compensation)
        {
            // Probe 12
            var pm = probeMeasurements.Values.FirstOrDefault(p => p.ProbeId == "Probe 12");
            double PM2 = dbRefDict.TryGetValue("RN-4", out var refValue1) ? refValue1.Min : 0.0;


            if (pm?.Readings.Any() != true) return double.NaN;

            double MaxReading = pm.Readings.Min();

            double runout = MaxReading - PM2;

            if (runout < 0)
            {
                runout = 0.001;
            }
            if (mode == ProcedureMode.Measurement && compensation != 0) runout += compensation;
            return Math.Abs(Math.Round(runout, 3));

        }

        private double CalculateRN5(Dictionary<string, ProbeMeasurement> probeMeasurements, Dictionary<string, (double Min, double Max)> dbRefDict, ProcedureMode mode, int signChange, double compensation)
        {
            // Probe 14
            var pm = probeMeasurements.Values.FirstOrDefault(p => p.ProbeId == "Probe 14");
            double PM2 = dbRefDict.TryGetValue("RN-5", out var refValue1) ? refValue1.Min : 0.0;


            if (pm?.Readings.Any() != true) return double.NaN;

            double MaxReading = pm.Readings.Min();

            double runout = MaxReading - PM2;

            if (runout < 0)
            {
                runout = 0.001;
            }
            if (mode == ProcedureMode.Measurement && compensation != 0) runout += compensation;
            return Math.Abs(Math.Round(runout, 3));

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
                // ✅ CHANGE: Use ParameterName as key (not ProbeId)
                string parameterName = probe.ParameterName;  // "Head Diameter"

                var config = masterVals?.FirstOrDefault(m =>
                    string.Equals(m.Parameter, parameterName, StringComparison.OrdinalIgnoreCase));

                var pm = new ProbeMeasurement
                {
                    ProbeId = parameterName,        // ✅ "Head Diameter"
                    MasterValue = config?.Nominal ?? 0,
                    TolerancePlus = config?.RTolPlus ?? 0,
                    ToleranceMinus = config?.RTolMinus ?? 0,
                    SignChange = config?.Sign_Change ?? 0,
                    Compensation = config?.Compensation ?? 0
                };

                _probeMeasurements[parameterName] = pm;  // ✅ Key = "Head Diameter"
                _orderedProbeMeasurements.Add(pm);
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
                var config = masterVals?
                    .FirstOrDefault(m =>
                        string.Equals(m.Para_No, probe.ProbeName, StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(m.Parameter, probe.ProbeName, StringComparison.OrdinalIgnoreCase)
                    );

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

            System.Diagnostics.Debug.WriteLine($"[DEBUG] Loaded {_probeMeasurements.Count} probe configs for part {partCode}");
            foreach (var pm in _orderedProbeMeasurements)
            {
                System.Diagnostics.Debug.WriteLine($"  Probe {pm.ProbeId} ({pm.Name}): Master={pm.MasterValue}, Tol±={pm.TolerancePlus}/{pm.ToleranceMinus}");
            }

            await StartMeasurementProcessAsync();
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
