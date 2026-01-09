using Microsoft.Data.SqlClient;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Configuration;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;

namespace EVMS
{
    public partial class ProbeInstallPage : UserControl, INotifyPropertyChanged
    {
        // 🔥 predefined probe names
        public static readonly string[] ProbeNamesStatic = {
                "Probe 1", "Probe 2", "Probe 3", "Probe 4", "Probe 5",
                "Probe 6", "Probe 7", "Probe 8", "Probe 9", "Probe 10",
                "Probe 11", "Probe 12", "Probe 13", "Probe 14", "Probe 15",
                "Probe 16", "Probe 17", "Probe 18", "Probe 19", "Probe 20",
                "Probe 21", "Probe 22", "Probe 23"
            };


        public static readonly string[] ProbeTypeStatic = {
            "Normal", "Pneumatic"
        };
        private readonly string connectionString;

        // existing collection of parameters (loaded from DB)
        public ObservableCollection<ParameterItem> Parameters { get; set; } = new ObservableCollection<ParameterItem>();

        // UI helper lists for assignment panel
        public ObservableCollection<int> AvailableBoxes { get; set; } = new ObservableCollection<int>(Enumerable.Range(1, 10));
        public ObservableCollection<int> AvailableChannels { get; set; } = new ObservableCollection<int>(Enumerable.Range(1, 4));
        public ObservableCollection<string> ProbeNames { get; set; } = new ObservableCollection<string>(ProbeNamesStatic);
        public ObservableCollection<string> ProbeType { get; set; } = new ObservableCollection<string>(ProbeTypeStatic);


        // Selected items bound to UI
        private ParameterItem? _selectedParameter;
        public ParameterItem? SelectedParameter
        {
            get => _selectedParameter;
            set { _selectedParameter = value; NotifyPropertyChanged(); }
        }

        private int _selectedBox;
        public int SelectedBox
        {
            get => _selectedBox;
            set { _selectedBox = value; NotifyPropertyChanged(); }
        }

        private int _selectedChannel;
        public int SelectedChannel
        {
            get => _selectedChannel;
            set { _selectedChannel = value; NotifyPropertyChanged(); }
        }

        private string? _selectedProbeName;
        public string? SelectedProbeName
        {
            get => _selectedProbeName;
            set { _selectedProbeName = value; NotifyPropertyChanged(); }
        }

        private string? _selectedProbeType;
        public string? SelectedProbeType
        {
            get => _selectedProbeType;
            set { _selectedProbeType = value; NotifyPropertyChanged(); }
        }

        // collection of mappings shown in DataGrid
        public ObservableCollection<ProbeMapping> Mappings { get; set; } = new ObservableCollection<ProbeMapping>();

        public ProbeInstallPage()
        {
            InitializeComponent();
            connectionString = ConfigurationManager.ConnectionStrings["EVMSDb"].ConnectionString!;
            DataContext = this;

            Loaded += async (s, e) => await InitializePageAsync();
        }

        private async Task InitializePageAsync()
        {
            await LoadParametersAsync();
            await LoadMappingsFromDbAsync();

            // default selections
            if (AvailableBoxes.Any()) SelectedBox = AvailableBoxes.First();
            if (AvailableChannels.Any()) SelectedChannel = AvailableChannels.First();
            SelectedProbeName = ProbeNames.FirstOrDefault();
            SelectedProbeType = ProbeType.FirstOrDefault();

        }

        // ------------------ LOAD PARAMETERS ------------------  
        private async Task LoadParametersAsync()
        {
            Parameters.Clear();

            try
            {
                using var con = new SqlConnection(connectionString);
                await con.OpenAsync();

                string partQuery = "SELECT TOP 1 Para_No FROM Part_Entry WHERE ActivePart = 1";
                object? raw = await new SqlCommand(partQuery, con).ExecuteScalarAsync();

                if (raw == null || raw == DBNull.Value)
                {
                    return;
                }

                string partNo = Convert.ToString(raw, CultureInfo.InvariantCulture) ?? "";
                if (string.IsNullOrWhiteSpace(partNo))
                {
                    return;
                }

                // Load parameters FIRST
                string paramQuery = "SELECT Parameter FROM PartConfig WHERE Para_No=@P AND ProbeStatus='Probe'";
                using var cmd = new SqlCommand(paramQuery, con);
                cmd.Parameters.AddWithValue("@P", partNo);

                using var r = await cmd.ExecuteReaderAsync();
                while (await r.ReadAsync())
                    Parameters.Add(new ParameterItem(r.GetString(0)));
                r.Close();

            }
            catch (Exception ex)
            {
                MessageBox.Show("Load error: " + ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // ------------------ LOAD EXISTING MAPPINGS ------------------
        private async Task LoadMappingsFromDbAsync()
        {
            Mappings.Clear();
            try
            {
                using var con = new SqlConnection(connectionString);
                await con.OpenAsync();

                string partQuery = "SELECT TOP 1 Para_No FROM Part_Entry WHERE ActivePart = 1";
                string partNo = Convert.ToString(await new SqlCommand(partQuery, con).ExecuteScalarAsync())!;

                if (partNo == null) return;

                string mapQuery = @"SELECT ParameterName, BoxId, ChannelId,ISNULL(ProbeName,'') as ProbeName ,ISNULL(ProbeType,'') as ProbeType
                                    FROM ProbeInstallationData WHERE PartNo=@P";

                using var mcmd = new SqlCommand(mapQuery, con);
                mcmd.Parameters.AddWithValue("@P", partNo);

                using var reader = await mcmd.ExecuteReaderAsync();
                while (await reader.ReadAsync())
                {
                    var mapping = new ProbeMapping
                    {
                        Parameter = reader.GetString(0),
                        BoxId = reader.GetInt32(1),
                        ChannelId = reader.GetInt32(2),
                        ProbeName = reader.IsDBNull(3) ? "" : reader.GetString(3),
                        ProbeType = reader.GetString(4),

                    };
                    Mappings.Add(mapping);
                }
            }
            catch (Exception ex)
            {
                // non-fatal: show status
            }
        }

        // ------------------ ADD MAPPING ------------------
        private void BtnAddMapping_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (SelectedParameter == null)
                {
                    MessageBox.Show("Please select a parameter.", "Missing", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                if (SelectedBox < 1 || SelectedBox > 10)
                {
                    MessageBox.Show("Please select a valid box (1-10).", "Invalid", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                if (SelectedChannel < 1 || SelectedChannel > 4)
                {
                    MessageBox.Show("Please select a valid channel (1-4).", "Invalid", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                if (string.IsNullOrWhiteSpace(SelectedProbeName))
                {
                    MessageBox.Show("Please select a probe name.", "Missing", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                if (string.IsNullOrWhiteSpace(SelectedProbeType))
                {
                    MessageBox.Show("Please select a probe name.", "Missing", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                var map = new ProbeMapping
                {
                    Parameter = SelectedParameter.Name,
                    BoxId = SelectedBox,
                    ChannelId = SelectedChannel,
                    ProbeName = SelectedProbeName ?? "",
                    ProbeType = SelectedProbeType ?? ""

                };

                Mappings.Add(map);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Add mapping error: " + ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // ------------------ DELETE MAPPING ------------------
        private void BtnDeleteMapping_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.DataContext is ProbeMapping map)
            {
                if (MessageBox.Show($"Delete mapping {map.Parameter} Box {map.BoxId} Ch {map.ChannelId} ?", "Confirm", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                {
                    Mappings.Remove(map);
                }
            }
        }

        // ------------------ SAVE ALL MAPPINGS TO DB ------------------
        private async void BtnSaveMappingsToDb_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                using var con = new SqlConnection(connectionString);
                await con.OpenAsync();

                // get current active part
                string partQuery = "SELECT TOP 1 Para_No FROM Part_Entry WHERE ActivePart = 1";
                string partNo = Convert.ToString(await new SqlCommand(partQuery, con).ExecuteScalarAsync())!;

                if (string.IsNullOrWhiteSpace(partNo))
                {
                    MessageBox.Show("No active part selected.", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                // delete existing mappings for this part
                var del = new SqlCommand("DELETE FROM ProbeInstallationData WHERE PartNo=@P", con);
                del.Parameters.AddWithValue("@P", partNo);
                await del.ExecuteNonQueryAsync();

                int saved = 0;
                foreach (var m in Mappings)
                {
                    var ins = new SqlCommand(
                        "INSERT INTO ProbeInstallationData (PartNo,ParameterName,BoxId,ChannelId,ProbeName,ProbeType) VALUES (@p,@n,@b,@c,@probe,@probetype)", con);
                    ins.Parameters.AddWithValue("@p", partNo);
                    ins.Parameters.AddWithValue("@n", m.Parameter);
                    ins.Parameters.AddWithValue("@b", m.BoxId);
                    ins.Parameters.AddWithValue("@c", m.ChannelId);
                    ins.Parameters.AddWithValue("@probe", m.ProbeName ?? "");
                    ins.Parameters.AddWithValue("@probetype", m.ProbeType ?? "");

                    await ins.ExecuteNonQueryAsync();
                    saved++;
                }

                MessageBox.Show($"Saved {saved} mappings.", "Saved", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Save error: " + ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // ------------------ LEGACY SAVE (kept for compatibility) ------------------
        private async void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                using var con = new SqlConnection(connectionString);
                await con.OpenAsync();

                string partQuery = "SELECT TOP 1 Para_No FROM Part_Entry WHERE ActivePart = 1";
                string partNo = Convert.ToString(await new SqlCommand(partQuery, con).ExecuteScalarAsync())!;

                if (partNo == null)
                {
                    MessageBox.Show("No active part selected.", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                // Delete existing mappings for this part
                var del = new SqlCommand("DELETE FROM ProbeInstallationData WHERE PartNo=@P", con);
                del.Parameters.AddWithValue("@P", partNo);
                await del.ExecuteNonQueryAsync();

                int savedCount = 0;
                foreach (var p in Parameters.Where(x => x.BoxId > 0 && x.Channels.Any()))
                {
                    foreach (int ch in p.Channels.Where(c => c >= 1 && c <= 4))
                    {
                        var ins = new SqlCommand(
                            "INSERT INTO ProbeInstallationData (PartNo,ParameterName,BoxId,ChannelId,ProbeName,ProbeType) VALUES (@p,@n,@b,@c,@probeName,@ProbeType)", con);
                        ins.Parameters.AddWithValue("@p", partNo);
                        ins.Parameters.AddWithValue("@n", p.Name);
                        ins.Parameters.AddWithValue("@b", p.BoxId);
                        ins.Parameters.AddWithValue("@c", ch);
                        ins.Parameters.AddWithValue("@probeName", p.ProbeName ?? "");
                        ins.Parameters.AddWithValue("@ProbeType", p.ProbeType ?? "");

                        await ins.ExecuteNonQueryAsync();
                        savedCount++;
                    }
                }

                MessageBox.Show($"🎉 Probe assignments saved successfully!\n{savedCount} mappings stored for {Parameters.Count(p => p.Channels.Any())} parameters.", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Save error: " + ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // INotifyPropertyChanged implementation (for bindings)
        public event PropertyChangedEventHandler? PropertyChanged;
        private void NotifyPropertyChanged([CallerMemberName] string propName = "") =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propName));
    }

    // ------------------ models ------------------
    public class ParameterItem : INotifyPropertyChanged
    {
        private string _name = "";
        private int _boxId;
        private string _liveValue = "-";
        private string _probeName = "";
        private string _probeType = "";


        public string Name
        {
            get => _name;
            set { _name = value; Notify(); }
        }

        public int BoxId
        {
            get => _boxId;
            set { _boxId = value; Notify(); }
        }

        public ObservableCollection<int> Channels { get; set; } = new ObservableCollection<int>();

        public string ProbeName
        {
            get => _probeName;
            set { _probeName = value; Notify(); }
        }

        public string ProbeType
        {
            get => _probeType;
            set { _probeType = value; Notify(); }
        }

        public string ChannelsAsString
        {
            get => string.Join(",", Channels);
            set
            {
                Channels.Clear();
                if (!string.IsNullOrWhiteSpace(value))
                    foreach (var p in value.Split(',', ';', ' '))
                        if (int.TryParse(p.Trim(), out int ch) && ch >= 1 && ch <= 4)
                            Channels.Add(ch);
                Notify();
            }
        }

        public string LiveValue
        {
            get => _liveValue;
            set { _liveValue = value; Notify(); }
        }

        public ParameterItem(string name) => Name = name;
        public ParameterItem() { }

        public event PropertyChangedEventHandler? PropertyChanged;
        private void Notify([CallerMemberName] string prop = "") =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(prop));
    }

    public class ProbeMapping
    {
        public string Parameter { get; set; } = "";
        public int BoxId { get; set; }
        public int ChannelId { get; set; }
        public string ProbeName { get; set; } = "";
        public string ProbeType { get; set; } = "";
    }
}
