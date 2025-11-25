using Microsoft.Data.SqlClient;
using System;
using System.Configuration;
using System.Data;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace EVMS
{
    public partial class AdminControlePage : UserControl
    {
        private readonly string connectionString;
        public event Action<string>? StatusMessageChanged;

        public AdminControlePage()
        {
            InitializeComponent();
            

            connectionString = ConfigurationManager.ConnectionStrings["EVMSDb"].ConnectionString;

            EnsureUsersTableExists();  // Create table if missing

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

        private void UpdateStatus(string message)
        {
            StatusMessageChanged?.Invoke(message);
        }

        private void EnsureUsersTableExists()
        {
            string createTableQuery = @"
            IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='Users' and xtype='U')
            CREATE TABLE Users
            (
                ID INT IDENTITY(1,1) PRIMARY KEY,
                UserType NVARCHAR(50) NOT NULL,
                UserID NVARCHAR(50) NOT NULL UNIQUE,
                UserName NVARCHAR(100) NOT NULL,
                Password NVARCHAR(100) NOT NULL
            )";

            using (SqlConnection con = new SqlConnection(connectionString))
            {
                con.Open();
                using (SqlCommand cmd = new SqlCommand(createTableQuery, con))
                {
                    cmd.ExecuteNonQuery();
                }
            }
        }

        private void BtnAdd_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string userType = (cmbPartNo.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "";
                string userID = txtID.Text.Trim();
                string userName = txtName.Text.Trim();
                string role = txtRTolMinus.Text.Trim();

                if (string.IsNullOrWhiteSpace(userID) || string.IsNullOrWhiteSpace(userName) || string.IsNullOrWhiteSpace(role))
                {
                    MessageBox.Show("UserID, Name and Role cannot be empty.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                using (SqlConnection con = new SqlConnection(connectionString))
                {
                    con.Open();

                    // Check if User ID already exists
                    string checkQuery = "SELECT COUNT(*) FROM Users WHERE UserID = @UserID";
                    using (SqlCommand checkCmd = new SqlCommand(checkQuery, con))
                    {
                        checkCmd.Parameters.AddWithValue("@UserID", userID);
                        int count = (int)checkCmd.ExecuteScalar();

                        if (count > 0)
                        {
                            MessageBox.Show("User ID already exists.", "Duplicate Entry", MessageBoxButton.OK, MessageBoxImage.Warning);
                            return;
                        }
                    }

                    string insertQuery = "INSERT INTO Users (UserType, UserID, UserName, Password) VALUES (@UserType, @UserID, @UserName, @Password)";
                    using (SqlCommand cmd = new SqlCommand(insertQuery, con))
                    {
                        cmd.Parameters.AddWithValue("@UserType", userType);
                        cmd.Parameters.AddWithValue("@UserID", userID);
                        cmd.Parameters.AddWithValue("@UserName", userName);
                        cmd.Parameters.AddWithValue("@Password", role);
                        cmd.ExecuteNonQuery();
                    }
                }

                UpdateStatus("User Added Sucessesfuly...!!");
                LoadData();
                ClearInputs();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BtnUpdate_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (dataGrid.SelectedItem is not DataRowView selectedRow)
                {
                    MessageBox.Show("Please select a user to update.", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                int id = Convert.ToInt32(selectedRow["ID"]);
                string userType = (cmbPartNo.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "";
                string userID = txtID.Text.Trim();
                string userName = txtName.Text.Trim();
                string role = txtRTolMinus.Text.Trim();

                if (string.IsNullOrWhiteSpace(userID) || string.IsNullOrWhiteSpace(userName) || string.IsNullOrWhiteSpace(role))
                {
                    MessageBox.Show("UserID, Name and Role cannot be empty.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                using (SqlConnection con = new SqlConnection(connectionString))
                {
                    con.Open();

                    // Check if UserID exists in another record
                    string checkQuery = "SELECT COUNT(*) FROM Users WHERE UserID = @UserID AND ID != @ID";
                    using (SqlCommand checkCmd = new SqlCommand(checkQuery, con))
                    {
                        checkCmd.Parameters.AddWithValue("@UserID", userID);
                        checkCmd.Parameters.AddWithValue("@ID", id);
                        int count = (int)checkCmd.ExecuteScalar();

                        if (count > 0)
                        {
                            MessageBox.Show("User ID already exists in another record.", "Duplicate Entry", MessageBoxButton.OK, MessageBoxImage.Warning);
                            return;
                        }
                    }

                    string updateQuery = "UPDATE Users SET UserType = @UserType, UserID = @UserID, UserName = @UserName, Password = @Password WHERE ID = @ID";
                    using (SqlCommand cmd = new SqlCommand(updateQuery, con))
                    {
                        cmd.Parameters.AddWithValue("@UserType", userType);
                        cmd.Parameters.AddWithValue("@UserID", userID);
                        cmd.Parameters.AddWithValue("@UserName", userName);
                        cmd.Parameters.AddWithValue("@Password", role);
                        cmd.Parameters.AddWithValue("@ID", id);
                        cmd.ExecuteNonQuery();
                    }
                }

                MessageBox.Show("User updated successfully.", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                LoadData();
                ClearInputs();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BtnDelete_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (dataGrid.SelectedItem is not DataRowView selectedRow)
                {
                    MessageBox.Show("Please select a user to delete.", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                int id = Convert.ToInt32(selectedRow["ID"]);
                string userName = selectedRow["UserName"].ToString();

                if (MessageBox.Show($"Are you sure you want to delete user '{userName}'?", "Confirm Delete",
                    MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.No)
                    return;

                using (SqlConnection con = new SqlConnection(connectionString))
                {
                    con.Open();

                    string deleteQuery = "DELETE FROM Users WHERE ID = @ID";
                    using (SqlCommand cmd = new SqlCommand(deleteQuery, con))
                    {
                        cmd.Parameters.AddWithValue("@ID", id);
                        cmd.ExecuteNonQuery();
                    }
                }

                UpdateStatus("User Deleted Sucessesfuly...!!");
                LoadData();
                ClearInputs();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

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
                    UserType,
                    UserID,
                    UserName,
                    Password
                FROM Users
                ORDER BY ID";

                    SqlDataAdapter da = new SqlDataAdapter(query, con);
                    DataTable dt = new DataTable();
                    da.Fill(dt);

                    // Add SrNo column if not present
                    if (!dt.Columns.Contains("SrNo"))
                    {
                        dt.Columns.Add("SrNo", typeof(int));
                    }

                    // Assign serial number based on row index
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        dt.Rows[i]["SrNo"] = i + 1;
                    }

                    dataGrid.ItemsSource = dt.DefaultView;

                    // Hide ID column if you want
                    if (dataGrid.Columns.Count > 0)
                        dataGrid.Columns[0].Visibility = Visibility.Collapsed;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading data: {ex.Message}", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }



        private void ClearInputs()
        {
            txtID.Text = "";
            txtName.Text = "";
            txtRTolMinus.Text = "";
            cmbPartNo.SelectedIndex = -1;

            btnUpdate.IsEnabled = false;
            btnDelete.IsEnabled = false;
            dataGrid.UnselectAll();
        }

        private void dataGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (dataGrid.SelectedItem is DataRowView row)
            {
                cmbPartNo.SelectedItem = null;
                foreach (ComboBoxItem item in cmbPartNo.Items)
                {
                    if (item.Content.ToString() == row["UserType"].ToString())
                    {
                        cmbPartNo.SelectedItem = item;
                        break;
                    }
                }

                txtID.Text = row["UserID"]?.ToString();
                txtName.Text = row["UserName"]?.ToString();
                txtRTolMinus.Text = row["Password"]?.ToString();

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
