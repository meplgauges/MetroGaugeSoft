using Microsoft.Data.SqlClient;
using System.Configuration;
using System.Data;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media.Imaging;

namespace EVMS
{
    public partial class AdminControlePage : UserControl
    {
        private readonly string connectionString;
        public event Action<string>? StatusMessageChanged;


        private string? selectedLogoFileName = null;

        public AdminControlePage()
        {
            InitializeComponent();

            connectionString = ConfigurationManager.ConnectionStrings["EVMSDb"].ConnectionString;

            EnsureUsersTableExists();

            btnAdd.Click += BtnAdd_Click;
            btnUpdate.Click += BtnUpdate_Click;
            btnDelete.Click += BtnDelete_Click;

            btnUpdate.IsEnabled = false;
            btnDelete.IsEnabled = false;

            this.Loaded += SettingsPage_Loaded;
            this.PreviewKeyDown += SettingsPage_PreviewKeyDown;

            LoadData();
            LoadCompanyInfo();
            ClearInputs();
        }

        // =============================================================
        // GENERAL
        // =============================================================

        private void SettingsPage_Loaded(object? sender, RoutedEventArgs e)
        {
            this.Focusable = true;
            this.IsTabStop = true;
            Keyboard.Focus(this);
        }

        private void SettingsPage_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                HandleEscKeyAction();
                e.Handled = true;
            }
        }

        private void HandleEscKeyAction()
        {
            Window currentWindow = Window.GetWindow(this);
            if (currentWindow != null)
            {
                var mainContentGrid = currentWindow.FindName("MainContentGrid") as Grid;
                if (mainContentGrid != null)
                {
                    mainContentGrid.Children.Clear();
                    mainContentGrid.Children.Add(new Dashboard());
                }
            }
        }

        private void UpdateStatus(string message)
        {
            StatusMessageChanged?.Invoke(message);
        }

        // =============================================================
        // DATABASE TABLE HANDLING
        // =============================================================

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

            using SqlConnection con = new SqlConnection(connectionString);
            con.Open();
            using SqlCommand cmd = new SqlCommand(createTableQuery, con);
            cmd.ExecuteNonQuery();
        }

        // =============================================================
        // USER CRUD
        // =============================================================

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
                    MessageBox.Show("UserID, Name and Role cannot be empty.");
                    return;
                }

                using SqlConnection con = new SqlConnection(connectionString);
                con.Open();

                string checkQuery = "SELECT COUNT(*) FROM Users WHERE UserID = @UserID";
                using SqlCommand checkCmd = new SqlCommand(checkQuery, con);
                checkCmd.Parameters.AddWithValue("@UserID", userID);
                int count = (int)checkCmd.ExecuteScalar();
                if (count > 0)
                {
                    MessageBox.Show("User ID already exists.");
                    return;
                }

                string insertQuery = "INSERT INTO Users (UserType, UserID, UserName, Password) VALUES (@UserType, @UserID, @UserName, @Password)";
                using SqlCommand cmd = new SqlCommand(insertQuery, con);
                cmd.Parameters.AddWithValue("@UserType", userType);
                cmd.Parameters.AddWithValue("@UserID", userID);
                cmd.Parameters.AddWithValue("@UserName", userName);
                cmd.Parameters.AddWithValue("@Password", role);
                cmd.ExecuteNonQuery();

                UpdateStatus("User Added Successfully!");
                LoadData();
                ClearInputs();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void BtnUpdate_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (dataGrid.SelectedItem is not DataRowView selectedRow)
                {
                    MessageBox.Show("Select a user to update.");
                    return;
                }

                int id = Convert.ToInt32(selectedRow["ID"]);
                string userType = (cmbPartNo.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "";
                string userID = txtID.Text.Trim();
                string userName = txtName.Text.Trim();
                string role = txtRTolMinus.Text.Trim();

                if (string.IsNullOrWhiteSpace(userID) || string.IsNullOrWhiteSpace(userName) || string.IsNullOrWhiteSpace(role))
                {
                    MessageBox.Show("UserID, Name and Role cannot be empty.");
                    return;
                }

                using SqlConnection con = new SqlConnection(connectionString);
                con.Open();

                string checkQuery = "SELECT COUNT(*) FROM Users WHERE UserID = @UserID AND ID != @ID";
                using SqlCommand checkCmd = new SqlCommand(checkQuery, con);
                checkCmd.Parameters.AddWithValue("@UserID", userID);
                checkCmd.Parameters.AddWithValue("@ID", id);

                if ((int)checkCmd.ExecuteScalar() > 0)
                {
                    MessageBox.Show("User ID exists in another record.");
                    return;
                }

                string updateQuery = "UPDATE Users SET UserType=@UserType, UserID=@UserID, UserName=@UserName, Password=@Password WHERE ID=@ID";
                using SqlCommand cmd = new SqlCommand(updateQuery, con);
                cmd.Parameters.AddWithValue("@UserType", userType);
                cmd.Parameters.AddWithValue("@UserID", userID);
                cmd.Parameters.AddWithValue("@UserName", userName);
                cmd.Parameters.AddWithValue("@Password", role);
                cmd.Parameters.AddWithValue("@ID", id);
                cmd.ExecuteNonQuery();

                MessageBox.Show("User updated successfully.");
                LoadData();
                ClearInputs();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void BtnDelete_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (dataGrid.SelectedItem is not DataRowView selectedRow)
                {
                    MessageBox.Show("Select a user to delete.");
                    return;
                }

                int id = Convert.ToInt32(selectedRow["ID"]);

                if (MessageBox.Show("Are you sure?", "Confirm", MessageBoxButton.YesNo) == MessageBoxResult.No)
                    return;

                using SqlConnection con = new SqlConnection(connectionString);
                con.Open();

                using SqlCommand cmd = new SqlCommand("DELETE FROM Users WHERE ID=@ID", con);
                cmd.Parameters.AddWithValue("@ID", id);
                cmd.ExecuteNonQuery();

                UpdateStatus("User Deleted Successfully!");
                LoadData();
                ClearInputs();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        // =============================================================
        // LOAD DATA
        // =============================================================

        private void LoadData()
        {
            try
            {
                using SqlConnection con = new SqlConnection(connectionString);
                con.Open();

                string query = "SELECT ID, UserType, UserID, UserName, Password FROM Users ORDER BY ID";
                SqlDataAdapter da = new SqlDataAdapter(query, con);
                DataTable dt = new DataTable();
                da.Fill(dt);

                if (!dt.Columns.Contains("SrNo"))
                    dt.Columns.Add("SrNo", typeof(int));

                for (int i = 0; i < dt.Rows.Count; i++)
                    dt.Rows[i]["SrNo"] = i + 1;

                dataGrid.ItemsSource = dt.DefaultView;

                if (dataGrid.Columns.Count > 0)
                    dataGrid.Columns[0].Visibility = Visibility.Collapsed;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        // =============================================================
        // INPUTS
        // =============================================================

        private void ClearInputs()
        {
            txtID.Clear();
            txtName.Clear();
            txtRTolMinus.Clear();
            cmbPartNo.SelectedIndex = -1;
            btnUpdate.IsEnabled = false;
            btnDelete.IsEnabled = false;
            dataGrid.UnselectAll();
        }

        private void dataGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (dataGrid.SelectedItem is DataRowView row)
            {
                foreach (ComboBoxItem item in cmbPartNo.Items)
                {
                    if (item.Content.ToString() == row["UserType"].ToString())
                    {
                        cmbPartNo.SelectedItem = item;
                        break;
                    }
                }

                txtID.Text = row["UserID"].ToString();
                txtName.Text = row["UserName"].ToString();
                txtRTolMinus.Text = row["Password"].ToString();

                btnUpdate.IsEnabled = true;
                btnDelete.IsEnabled = true;
            }
            else
            {
                ClearInputs();
            }
        }

        // =============================================================
        // IMAGE DIRECTORY FIX (REAL FIX)
        // =============================================================

        private string GetWritableLogoFolder()
        {
            string baseFolder = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            string logoFolder = Path.Combine(baseFolder, "EVMS", "CompanyLogo");

            if (!Directory.Exists(logoFolder))
                Directory.CreateDirectory(logoFolder);

            return logoFolder;
        }

        private void btnPickImage_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.Filter = "Image Files|*.png;*.jpg;*.jpeg;*.bmp";

            if (dlg.ShowDialog() == true)
            {
                string sourcePath = dlg.FileName;
                selectedLogoFileName = Path.GetFileName(sourcePath);

                string logoFolder = GetWritableLogoFolder();
                string destPath = Path.Combine(logoFolder, selectedLogoFileName);

                File.Copy(sourcePath, destPath, true);

                imgPreview.Source = new BitmapImage(new Uri(destPath));
            }
        }

        // =============================================================
        // COMPANY CONFIG
        // =============================================================

        private void btnCompanyUpdate_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string companyName = txtCompanyName.Text.Trim();

                if (string.IsNullOrEmpty(companyName))
                {
                    MessageBox.Show("Please enter company name.");
                    return;
                }

                string query = @"UPDATE CompanyConfig 
                             SET CompanyName=@name, LogoPath=@logo, UpdatedOn=GETDATE() 
                             WHERE Id=1";

                using SqlConnection conn = new SqlConnection(connectionString);
                using SqlCommand cmd = new SqlCommand(query, conn);

                cmd.Parameters.AddWithValue("@name", companyName);
                cmd.Parameters.AddWithValue("@logo", selectedLogoFileName ?? (object)DBNull.Value);

                conn.Open();
                cmd.ExecuteNonQuery();

                MessageBox.Show("Company updated successfully.");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void LoadCompanyInfo()
        {
            try
            {
                string query = "SELECT TOP 1 CompanyName, LogoPath FROM CompanyConfig WHERE Id=1";

                using SqlConnection conn = new SqlConnection(connectionString);
                using SqlCommand cmd = new SqlCommand(query, conn);

                conn.Open();
                using SqlDataReader reader = cmd.ExecuteReader();

                if (reader.Read())
                {
                    txtCompanyName.Text = reader["CompanyName"].ToString();
                    string logoFile = reader["LogoPath"]?.ToString();

                    if (!string.IsNullOrEmpty(logoFile))
                    {
                        string logoFolder = GetWritableLogoFolder();
                        string logoPath = Path.Combine(logoFolder, logoFile);

                        if (File.Exists(logoPath))
                        {
                            imgPreview.Source = new BitmapImage(new Uri(logoPath));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }

}
