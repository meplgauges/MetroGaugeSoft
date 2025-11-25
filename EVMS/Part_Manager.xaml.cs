using Microsoft.Data.SqlClient;
using System;
using System.Configuration;
using System.Data;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace EVMS
{
    public partial class Part_Manager : UserControl
    {
        private readonly string connectionString;

        public Part_Manager()
        {
            InitializeComponent();
            connectionString = ConfigurationManager.ConnectionStrings["EVMSDb"].ConnectionString;

            btnAdd.Click += BtnAdd_Click;
            btnUpdate.Click += BtnUpdate_Click;
            btnDelete.Click += BtnDelete_Click;

            btnUpdate.IsEnabled = false;
            btnDelete.IsEnabled = false;

            this.Loaded += SettingsPage_Loaded;

            // ✅ Register ESC key handler
            this.PreviewKeyDown += SettingsPage_PreviewKeyDown;
            LoadData();
            ClearInputs();
        }


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
        // Add method with validation and uniqueness check
        private void BtnAdd_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string paraNo = txtPartNumber.Text.Trim();
                string paraName = txtPartName.Text.Trim();
                bool activePart = chkActivePart.IsChecked == true;

                if (string.IsNullOrEmpty(paraNo) || string.IsNullOrEmpty(paraName))
                {
                    MessageBox.Show("⚠️ Part Number and Part Name cannot be empty.", "Validation Error",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                using (SqlConnection con = new SqlConnection(connectionString))
                {
                    con.Open();

                    // Check if Part Number already exists
                    string checkQuery = "SELECT COUNT(*) FROM Part_Entry WHERE Para_No = @Para_No";
                    using (SqlCommand checkCmd = new SqlCommand(checkQuery, con))
                    {
                        checkCmd.Parameters.AddWithValue("@Para_No", paraNo);
                        int count = (int)checkCmd.ExecuteScalar();

                        if (count > 0)
                        {
                            MessageBox.Show("⚠️ Part Number already exists. Please use a different Part Number.", "Duplicate Entry",
                                MessageBoxButton.OK, MessageBoxImage.Warning);
                            return;
                        }
                    }

                    // If new part is to be active, reset other active parts
                    if (activePart)
                    {
                        string resetActiveQuery = "UPDATE Part_Entry SET ActivePart = 0 WHERE ActivePart = 1";
                        using (SqlCommand resetCmd = new SqlCommand(resetActiveQuery, con))
                        {
                            resetCmd.ExecuteNonQuery();
                        }
                    }

                    // Insert new record with ActivePart value
                    string insertQuery = "INSERT INTO Part_Entry (Para_No, Para_Name, ActivePart) VALUES (@Para_No, @Para_Name, @ActivePart)";
                    using (SqlCommand cmd = new SqlCommand(insertQuery, con))
                    {
                        cmd.Parameters.AddWithValue("@Para_No", paraNo);
                        cmd.Parameters.AddWithValue("@Para_Name", paraName);
                        cmd.Parameters.AddWithValue("@ActivePart", activePart ? 1 : 0);
                        cmd.ExecuteNonQuery();
                    }
                }

                MessageBox.Show("✅ Record inserted successfully.", "Success",
                    MessageBoxButton.OK, MessageBoxImage.Information);

                LoadData();
                ClearInputs();
            }
            catch (SqlException sqlEx)
            {
                MessageBox.Show($"Database error: {sqlEx.Message}", "SQL Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Unexpected error: {ex.Message}", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // Update method with validation and uniqueness check
        private void BtnUpdate_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (dataGrid.SelectedItem is not DataRowView selectedRow)
                {
                    MessageBox.Show("⚠️ Please select a record to update.", "Warning",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                string paraNo = txtPartNumber.Text.Trim();
                string paraName = txtPartName.Text.Trim();
                bool activePartUpdate = chkActivePart.IsChecked == true;

                if (string.IsNullOrEmpty(paraNo) || string.IsNullOrEmpty(paraName))
                {
                    MessageBox.Show("⚠️ Part Number and Part Name cannot be empty.", "Validation Error",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                int id = Convert.ToInt32(selectedRow["ID"]);

                using (SqlConnection con = new SqlConnection(connectionString))
                {
                    con.Open();

                    // Check if Part Number exists in another record
                    string checkQuery = "SELECT COUNT(*) FROM Part_Entry WHERE Para_No = @Para_No AND ID != @ID";
                    using (SqlCommand checkCmd = new SqlCommand(checkQuery, con))
                    {
                        checkCmd.Parameters.AddWithValue("@Para_No", paraNo);
                        checkCmd.Parameters.AddWithValue("@ID", id);
                        int count = (int)checkCmd.ExecuteScalar();

                        if (count > 0)
                        {
                            MessageBox.Show("⚠️ Part Number already exists in another record. Please use a different Part Number.", "Duplicate Entry",
                                MessageBoxButton.OK, MessageBoxImage.Warning);
                            return;
                        }
                    }

                    // If updating to active, reset others to inactive
                    if (activePartUpdate)
                    {
                        string resetActiveQuery = "UPDATE Part_Entry SET ActivePart = 0 WHERE ActivePart = 1 AND ID != @ID";
                        using (SqlCommand resetCmd = new SqlCommand(resetActiveQuery, con))
                        {
                            resetCmd.Parameters.AddWithValue("@ID", id);
                            resetCmd.ExecuteNonQuery();
                        }
                    }

                    // Update record with ActivePart value
                    string updateQuery = "UPDATE Part_Entry SET Para_No = @Para_No, Para_Name = @Para_Name, ActivePart = @ActivePart WHERE ID = @ID";
                    using (SqlCommand cmd = new SqlCommand(updateQuery, con))
                    {
                        cmd.Parameters.AddWithValue("@Para_No", paraNo);
                        cmd.Parameters.AddWithValue("@Para_Name", paraName);
                        cmd.Parameters.AddWithValue("@ActivePart", activePartUpdate ? 1 : 0);
                        cmd.Parameters.AddWithValue("@ID", id);

                        int rows = cmd.ExecuteNonQuery();
                        if (rows == 0)
                        {
                            MessageBox.Show("⚠️ No record found to update.", "Not Found",
                                MessageBoxButton.OK, MessageBoxImage.Warning);
                            return;
                        }
                    }
                }

                MessageBox.Show("✅ Record updated successfully.", "Success",
                    MessageBoxButton.OK, MessageBoxImage.Information);

                LoadData();
                ClearInputs();
            }
            catch (SqlException sqlEx)
            {
                MessageBox.Show($"Database error: {sqlEx.Message}", "SQL Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Unexpected error: {ex.Message}", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // Delete method
        private void BtnDelete_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (dataGrid.SelectedItem is not DataRowView row)
                {
                    MessageBox.Show("⚠️ Please select a record to delete.", "Warning",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                int id = Convert.ToInt32(row["ID"]);
                string? paraNo = row["Para_No"].ToString();

                if (MessageBox.Show($"Are you sure you want to delete '{paraNo}'?", "Confirm Delete",
                    MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.No)
                    return;

                using (SqlConnection con = new SqlConnection(connectionString))
                {
                    con.Open();

                    string query = "DELETE FROM Part_Entry WHERE ID = @ID";
                    using (SqlCommand cmd = new SqlCommand(query, con))
                    {
                        cmd.Parameters.AddWithValue("@ID", id);
                        cmd.ExecuteNonQuery();
                    }
                }

                MessageBox.Show("✅ Record deleted successfully.", "Success",
                    MessageBoxButton.OK, MessageBoxImage.Information);

                LoadData();
                ClearInputs();
            }
            catch (SqlException sqlEx)
            {
                MessageBox.Show($"Database error: {sqlEx.Message}", "SQL Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Unexpected error: {ex.Message}", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // Load data and hide ID column
        private void LoadData()
        {
            try
            {
                using (SqlConnection con = new SqlConnection(connectionString))
                {
                    con.Open();

                    string query = @"
                        SELECT
                            ID,
                            ROW_NUMBER() OVER (ORDER BY ID) AS SrNo,
                            Para_No,
                            Para_Name,
                            ActivePart
                        FROM Part_Entry
                        ORDER BY ID";

                    SqlDataAdapter da = new SqlDataAdapter(query, con);
                    DataTable dt = new DataTable();
                    da.Fill(dt);

                    dataGrid.ItemsSource = dt.DefaultView;

                    // Hide the ID column (first column)
                    if (dataGrid.Columns.Count > 0)
                        dataGrid.Columns[0].Visibility = Visibility.Collapsed;
                }
            }
            catch (SqlException sqlEx)
            {
                MessageBox.Show($"Database error: {sqlEx.Message}", "SQL Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Unexpected error: {ex.Message}", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // Clear input fields and disable buttons
        private void ClearInputs()
        {
            txtPartNumber.Text = string.Empty;
            txtPartName.Text = string.Empty;
            chkActivePart.IsChecked = false;
            btnUpdate.IsEnabled = false;
            btnDelete.IsEnabled = false;
        }

        // Handle DataGrid selection changed
        private void dataGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (dataGrid.SelectedItem is DataRowView row)
            {
                txtPartNumber.Text = row["Para_No"].ToString();
                txtPartName.Text = row["Para_Name"].ToString();
                chkActivePart.IsChecked = row["ActivePart"].ToString() == "1";
                btnUpdate.IsEnabled = true;
                btnDelete.IsEnabled = true;
            }
            else
            {
                ClearInputs();
            }
        }

    }
}
