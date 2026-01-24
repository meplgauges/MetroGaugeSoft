using EVMS.Service;
using Microsoft.Data.SqlClient;
using System.Configuration;
using System.Data;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace EVMS
{
    public partial class MasterReadingPage : UserControl
    {
        private readonly string connectionString;
        private DataStorageService dataStorageService;


        public MasterReadingPage()
        {
            InitializeComponent();
            connectionString = ConfigurationManager.ConnectionStrings["EVMSDb"].ConnectionString;
            btnAdd.Click += BtnAdd_Click;
            btnUpdate.Click += BtnUpdate_Click;
            btnDelete.Click += BtnDelete_Click;

            cmbPartNo.SelectionChanged += CmbPartNo_SelectionChanged;
            cmbParameter.SelectionChanged += CmbParameter_SelectionChanged;

            btnUpdate.IsEnabled = false;
            btnDelete.IsEnabled = false;
            dataStorageService = new DataStorageService();

            this.Loaded += SettingsPage_Loaded;

            // ✅ Register ESC key handler
            this.PreviewKeyDown += SettingsPage_PreviewKeyDown;
            LoadPartNumbers();
            LoadMasterExpirationData();  // add this line to load saved data in UI

        }

        // ✅ ESC key detection
        private void SettingsPage_PreviewKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Escape)
            {
                HandleEscKeyAction();
                e.Handled = true;
            }
        }


        private void SettingsPage_Loaded(object? sender, RoutedEventArgs e)
        {
            // Ask WPF to focus this control (deferred)
            this.Focusable = true;
            this.IsTabStop = true;

            // Try several ways to set keyboard focus
            Keyboard.Focus(this);                                  // set logical focus
            FocusManager.SetFocusedElement(Window.GetWindow(this)!, this); // set focused element on window
        }
        // ✅ Handles ESC key press to go back to HomePage
        private void HandleEscKeyAction()
        {
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
                }
            }
        }

        private void LoadMasterExpirationData()
        {
            try
            {
                var (mode, setValue, _) = dataStorageService.GetMasterExpiration();

                if (mode == 1) // Count mode
                {
                    rbCount.IsChecked = true;
                    lblInput.Text = "Set Count (1 - 9999)";
                    txtInput.MaxLength = 4;
                }
                else // Time mode
                {
                    rbTime.IsChecked = true;
                    lblInput.Text = "Set Time (1 - 24 Hours)";
                    txtInput.MaxLength = 2;
                }

                txtInput.Text = setValue.ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to load master expiration data: " + ex.Message);
            }
        }


        private void LoadPartNumbers()
        {
            try
            {
                using (SqlConnection con = new SqlConnection(connectionString))
                {
                    con.Open();
                    string query = "SELECT Para_No FROM PART_ENTRY  WHERE ActivePart = 1";
                    SqlCommand cmd = new SqlCommand(query, con);
                    SqlDataReader reader = cmd.ExecuteReader();

                    cmbPartNo.Items.Clear();
                    while (reader.Read())
                    {
                        cmbPartNo.Items.Add(reader["Para_No"].ToString());
                    }

                    if (cmbPartNo.Items.Count > 0)
                        cmbPartNo.SelectedIndex = 0;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading Part Numbers: {ex.Message}", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LoadParametersByParaNo(string paraNo)
        {
            try
            {
                using var con = new SqlConnection(connectionString);
                con.Open();
                using var cmd = new SqlCommand("SELECT DISTINCT Parameter FROM PartConfig WHERE Para_No = @Para_No", con);
                cmd.Parameters.AddWithValue("@Para_No", paraNo);
                using var reader = cmd.ExecuteReader();

                cmbParameter.Items.Clear();
                int count = 0;
                while (reader.Read())
                {
                    cmbParameter.Items.Add(reader["Parameter"].ToString());
                    count++;
                }
                if (count > 0)
                    cmbParameter.SelectedIndex = 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading Parameters: {ex.Message}", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LoadMasterReadingDataByParaNo(string paraNo)
        {
            try
            {
                using (SqlConnection con = new SqlConnection(connectionString))
                {
                    con.Open();

                    string query = @"
                        SELECT 
                            Id,
                            ROW_NUMBER() OVER (ORDER BY Id) AS RowNo,
                            Para_No,
                            Parameter, 
                            Nominal, 
                            RTolPlus, 
                            RTolMinus
                        FROM MasterReadingData
                        WHERE Para_No = @Para_No
                        ORDER BY Id";
                    SqlCommand cmd = new SqlCommand(query, con);
                    cmd.Parameters.AddWithValue("@Para_No", paraNo);
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    dataGrid.ItemsSource = dt.DefaultView;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading data for Part No '{paraNo}': {ex.Message}", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void CmbPartNo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (cmbPartNo.SelectedItem == null)
                return;

            string paraNo = cmbPartNo.SelectedItem.ToString() ?? string.Empty;

            LoadParametersByParaNo(paraNo);
            LoadMasterReadingDataByParaNo(paraNo);
            ClearInputs();
        }

        private void CmbParameter_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (cmbParameter.SelectedItem == null)
                return;

            // You can optionally load or filter data based on parameter here
            ClearInputs();
        }

        private void BtnAdd_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (!ValidateInputs())
                    return;

                string paraNo = cmbPartNo.SelectedItem?.ToString() ?? string.Empty;
                string parameter = cmbParameter.SelectedItem?.ToString() ?? string.Empty;

                decimal nominal = ParseDecimal(txtNominal.Text);
                decimal rTolPlus = ParseDecimal(txtRTolPlus.Text);
                decimal rTolMinus = ParseDecimal(txtRTolMinus.Text);

                using var con = new SqlConnection(connectionString);
                con.Open();

                string query = @"INSERT INTO MasterReadingData
                                (Para_No, Parameter, Nominal, RTolPlus, RTolMinus)
                                VALUES (@Para_No, @Parameter, @Nominal, @RTolPlus, @RTolMinus)";

                using var cmd = new SqlCommand(query, con);
                cmd.Parameters.AddWithValue("@Para_No", paraNo);
                cmd.Parameters.AddWithValue("@Parameter", parameter);
                cmd.Parameters.AddWithValue("@Nominal", nominal);
                cmd.Parameters.AddWithValue("@RTolPlus", rTolPlus);
                cmd.Parameters.AddWithValue("@RTolMinus", rTolMinus);

                cmd.ExecuteNonQuery();

                MessageBox.Show("✅ Record inserted successfully.", "Success", MessageBoxButton.OK, MessageBoxImage.Information);

                LoadMasterReadingDataByParaNo(paraNo);
                ClearInputs();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error inserting record: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BtnUpdate_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (!ValidateInputs())
                    return;

                if (dataGrid.SelectedItem is not DataRowView row)
                {
                    MessageBox.Show("⚠️ Please select a record to update.", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                int id = Convert.ToInt32(row["Id"]);
                string paraNo = cmbPartNo.SelectedItem?.ToString() ?? string.Empty;
                string parameter = cmbParameter.SelectedItem?.ToString() ?? string.Empty;

                decimal nominal = ParseDecimal(txtNominal.Text);
                decimal rTolPlus = ParseDecimal(txtRTolPlus.Text);
                decimal rTolMinus = ParseDecimal(txtRTolMinus.Text);

                using var con = new SqlConnection(connectionString);
                con.Open();

                string query = @"UPDATE MasterReadingData SET
                                Para_No = @Para_No,
                                Parameter = @Parameter,
                                Nominal = @Nominal,
                                RTolPlus = @RTolPlus,
                                RTolMinus = @RTolMinus
                                WHERE Id = @Id";

                using var cmd = new SqlCommand(query, con);
                cmd.Parameters.AddWithValue("@Para_No", paraNo);
                cmd.Parameters.AddWithValue("@Parameter", parameter);
                cmd.Parameters.AddWithValue("@Nominal", nominal);
                cmd.Parameters.AddWithValue("@RTolPlus", rTolPlus);
                cmd.Parameters.AddWithValue("@RTolMinus", rTolMinus);
                cmd.Parameters.AddWithValue("@Id", id);

                int rows = cmd.ExecuteNonQuery();

                if (rows == 0)
                {
                    MessageBox.Show("⚠️ No record found to update.", "Not Found", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                MessageBox.Show("✅ Record updated successfully.", "Success", MessageBoxButton.OK, MessageBoxImage.Information);

                LoadMasterReadingDataByParaNo(paraNo);
                ClearInputs();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error updating record: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BtnDelete_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (dataGrid.SelectedItem is not DataRowView row)
                {
                    MessageBox.Show("⚠️ Please select a record to delete.", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                int id = Convert.ToInt32(row["Id"]);
                string paraNo = cmbPartNo.SelectedItem?.ToString() ?? string.Empty;

                if (MessageBox.Show($"Are you sure you want to delete the selected record?", "Confirm Delete",
                    MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.No)
                    return;

                using var con = new SqlConnection(connectionString);
                con.Open();

                string query = "DELETE FROM MasterReadingData WHERE Id = @Id";

                using var cmd = new SqlCommand(query, con);
                cmd.Parameters.AddWithValue("@Id", id);
                cmd.ExecuteNonQuery();

                MessageBox.Show("✅ Record deleted successfully.", "Success", MessageBoxButton.OK, MessageBoxImage.Information);

                LoadMasterReadingDataByParaNo(paraNo);
                ClearInputs();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error deleting record: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ClearInputs()
        {
            txtNominal.Clear();
            txtRTolPlus.Clear();
            txtRTolMinus.Clear();

            btnUpdate.IsEnabled = false;
            btnDelete.IsEnabled = false;
        }

        private decimal ParseDecimal(string input)
        {
            return decimal.TryParse(input, out var value) ? value : 0;
        }

        private bool ValidateInputs()
        {
            if (!decimal.TryParse(txtNominal.Text, out _))
            {
                MessageBox.Show("⚠️ Nominal value must be a number.", "Validation Error",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return false;
            }
            return true;
        }

        private void dataGrid_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            e.Row.Header = (e.Row.GetIndex() + 1).ToString();
        }

        private void dataGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (dataGrid.SelectedItem is DataRowView row)
            {
                string paraNo = row["Para_No"].ToString();
                if (!string.IsNullOrEmpty(paraNo) && cmbPartNo.Items.Contains(paraNo))
                {
                    cmbPartNo.SelectedItem = paraNo;
                }

                string parameter = row["Parameter"].ToString();
                if (!string.IsNullOrEmpty(parameter) && cmbParameter.Items.Contains(parameter))
                {
                    cmbParameter.SelectedItem = parameter;
                }

                txtNominal.Text = row["Nominal"].ToString();
                txtRTolPlus.Text = row["RTolPlus"].ToString();
                txtRTolMinus.Text = row["RTolMinus"].ToString();

                btnUpdate.IsEnabled = true;
                btnDelete.IsEnabled = true;
            }
        }


        private void RadioButton_Checked(object sender, RoutedEventArgs e)
        {
            if (lblInput == null || txtInput == null) return; // avoid null crash

            if (rbCount.IsChecked == true)
            {
                lblInput.Text = "Set Count (1 - 9999)";
                txtInput.Text = "";
                txtInput.MaxLength = 4;
            }
            else if (rbTime.IsChecked == true)
            {
                lblInput.Text = "Set Time (1 - 24 Hours)";
                txtInput.Text = "";
                txtInput.MaxLength = 2;
            }
        }


        private void NumberValidationTextBox(object sender, TextCompositionEventArgs e)
        {
            e.Handled = !int.TryParse(e.Text, out _); // Allow only digits
        }

        private void btnSet_Click(object sender, RoutedEventArgs e)
        {

            bool? isCountChecked = rbCount.IsChecked;
            bool? isTimeChecked = rbTime.IsChecked;

            if (isCountChecked != true && isTimeChecked != true)
            {
                MessageBox.Show("Please select Count or Time mode.");
                return;
            }

            int value, mode;
            if (isCountChecked == true)
            {
                if (!int.TryParse(txtInput.Text, out value) || value < 1 || value > 9999)
                {
                    MessageBox.Show("Count must be 1 to 9999.");
                    return;
                }
                mode = 1;
            }
            else
            {
                if (!int.TryParse(txtInput.Text, out value) || value < 1 || value > 24)
                {
                    MessageBox.Show("Time must be 1 to 24 hours.");
                    return;
                }
                mode = 0;
            }

            try
            {
                using var con = new SqlConnection(connectionString);
                con.Open();

                string updateSql = @"
            UPDATE MasterExpiration
            SET Mode = @Mode, SetValue = @SetValue, UpdatedAt = GETDATE()
            WHERE Id = 1";

                using var cmd = new SqlCommand(updateSql, con);
                cmd.Parameters.AddWithValue("@Mode", mode);
                cmd.Parameters.AddWithValue("@SetValue", value);

                int rowsAffected = cmd.ExecuteNonQuery();
                if (rowsAffected == 0)
                {
                    // If row does not exist (maybe first run), insert it
                    string insertSql = "INSERT INTO MasterExpiration (Id, Mode, SetValue, UpdatedAt) VALUES (1, @Mode, @SetValue, GETDATE())";
                    using var insertCmd = new SqlCommand(insertSql, con);
                    insertCmd.Parameters.AddWithValue("@Mode", mode);
                    insertCmd.Parameters.AddWithValue("@SetValue", value);
                    insertCmd.ExecuteNonQuery();
                }

                string modeText = mode == 1 ? "Count" : "Time";
                string messageValue = mode == 1 ? value.ToString() : $"{value} hours";
                MessageBox.Show($"{modeText} set to {messageValue}");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error saving data: " + ex.Message);
            }
        }


    }
}
