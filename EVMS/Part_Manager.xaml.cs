using Microsoft.Data.SqlClient;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Shapes;  // ✅ ADDED FOR ELLIPSE

namespace EVMS
{
    public partial class Part_Manager : UserControl
    {
        private readonly string connectionString;
        private List<ComboItem> idItems;
        private List<ComboItem> botItems;

        public Part_Manager()
        {
            InitializeComponent();
            connectionString = ConfigurationManager.ConnectionStrings["EVMSDb"].ConnectionString;

            btnAdd.Click += BtnAdd_Click;
            btnUpdate.Click += BtnUpdate_Click;
            btnDelete.Click += BtnDelete_Click;

            btnUpdate.IsEnabled = false;
            btnDelete.IsEnabled = false;

            this.Loaded += Part_Manager_Loaded;
            this.PreviewKeyDown += Part_Manager_PreviewKeyDown;

            LoadStaticComboData();
            LoadData();
            ClearInputs();
        }

        private void Part_Manager_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
                NavigateToDashboard();
        }

        private void Part_Manager_Loaded(object sender, RoutedEventArgs e)
        {
            this.Focusable = true;
            this.IsTabStop = true;
            Keyboard.Focus(this);
        }

        private void NavigateToDashboard()
        {
            var currentWindow = Window.GetWindow(this);
            if (currentWindow != null)
            {
                if (currentWindow.FindName("MainContentGrid") is Grid mainGrid)
                {
                    mainGrid.Children.Clear();
                    mainGrid.Children.Add(new Dashboard
                    {
                        HorizontalAlignment = HorizontalAlignment.Stretch,
                        VerticalAlignment = VerticalAlignment.Stretch
                    });
                }
            }
        }

        private void LoadStaticComboData()
        {
            idItems = new List<ComboItem>
            {
                new ComboItem{ Value=0, Text="NOT REQ"},
                new ComboItem{ Value=1, Text="ID-9.508"},
                new ComboItem{ Value=2, Text="ID-12.642"},
                new ComboItem{ Value=3, Text="ID-8.072"},
            };

            cmbCategory.ItemsSource = idItems;
            cmbCategory.DisplayMemberPath = "Text";
            cmbCategory.SelectedValuePath = "Value";
            cmbCategory.SelectedIndex = 0;

            botItems = new List<ComboItem>
            {
                new ComboItem{ Value=0, Text="DIRECT"},
                new ComboItem{ Value=1, Text="CYC1"},
                new ComboItem{ Value=2, Text="CYC2"},
                new ComboItem{ Value=3, Text="CYC3"}
            };

            cmbVendor.ItemsSource = botItems;
            cmbVendor.DisplayMemberPath = "Text";
            cmbVendor.SelectedValuePath = "Value";
            cmbVendor.SelectedIndex = 0;
        }

        // ✅ FIXED: Click Handler for Ellipse (Wrapped in Grid)
        private void StatusEllipse_Click(object sender, MouseButtonEventArgs e)
        {
            e.Handled = true; // prevent row selection conflict

            if (sender is not Border border ||
                border.DataContext is not DataRowView row)
                return;

            int clickedId = Convert.ToInt32(row["ID"]);
            bool currentActive = Convert.ToBoolean(row["IsActive"]);

            // If already active → do nothing
            if (currentActive)
                return;


            // 🔔 Confirmation dialog
            var result = MessageBox.Show(
                $"Do you want to activate this part?\n\nPart No: {row["Para_No"]}",
                "Confirm Activation",
                MessageBoxButton.OKCancel,
                MessageBoxImage.Question);

            if (result != MessageBoxResult.OK)
                return;

            try
            {
                using SqlConnection con = new(connectionString);
                con.Open();

                using SqlTransaction tran = con.BeginTransaction();

                try
                {
                    // 🔴 Deactivate ALL parts
                    using SqlCommand resetAll = new(
                        "UPDATE Part_Entry SET ActivePart = 0",
                        con, tran);
                    resetAll.ExecuteNonQuery();

                    // 🟢 Activate clicked part
                    using SqlCommand activate = new(
                        "UPDATE Part_Entry SET ActivePart = 1 WHERE ID = @ID",
                        con, tran);
                    activate.Parameters.AddWithValue("@ID", clickedId);
                    activate.ExecuteNonQuery();

                    tran.Commit();
                }
                catch
                {
                    tran.Rollback();
                    throw;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Failed to update status.\n\n{ex.Message}",
                    "Error",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
                return;
            }

            // 🔄 Refresh grid → 1 GREEN, others RED
            LoadData();
        }



        private void BtnAdd_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                int idValue = Convert.ToInt32(cmbCategory.SelectedValue);
                int botValue = Convert.ToInt32(cmbVendor.SelectedValue);
                string paraNo = txtPartNumber.Text.Trim();
                string paraName = txtPartName.Text.Trim();
                bool activePart = chkActivePart.IsChecked == true;

                if (string.IsNullOrEmpty(paraNo) || string.IsNullOrEmpty(paraName))
                {
                    MessageBox.Show("Part Number and Part Name cannot be empty.", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                using SqlConnection con = new(connectionString);
                con.Open();

                using SqlCommand check = new("SELECT COUNT(*) FROM Part_Entry WHERE Para_No=@Para_No", con);
                check.Parameters.AddWithValue("@Para_No", paraNo);
                int count = (int)check.ExecuteScalar();
                if (count > 0)
                {
                    MessageBox.Show("Part Number already exists.", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                if (activePart)
                {
                    using SqlCommand reset = new("UPDATE Part_Entry SET ActivePart=0 WHERE ActivePart=1", con);
                    reset.ExecuteNonQuery();
                }

                using SqlCommand insert = new("INSERT INTO Part_Entry (Para_No,Para_Name,ActivePart,ID_Value,BOT_Value) VALUES(@Para_No,@Para_Name,@ActivePart,@ID_Value,@BOT_Value)", con);
                insert.Parameters.AddWithValue("@Para_No", paraNo);
                insert.Parameters.AddWithValue("@Para_Name", paraName);
                insert.Parameters.AddWithValue("@ActivePart", activePart ? 1 : 0);
                insert.Parameters.AddWithValue("@ID_Value", idValue);
                insert.Parameters.AddWithValue("@BOT_Value", botValue);
                insert.ExecuteNonQuery();

                MessageBox.Show("Record added successfully.", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
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
            if (dataGrid.SelectedItem is not DataRowView row) return;

            try
            {
                int id = Convert.ToInt32(row["ID"]);
                int idValue = (int)cmbCategory.SelectedValue;
                int botValue = (int)cmbVendor.SelectedValue;
                string paraNo = txtPartNumber.Text.Trim();
                string paraName = txtPartName.Text.Trim();
                bool activePart = chkActivePart.IsChecked == true;

                using SqlConnection con = new(connectionString);
                con.Open();

                using SqlCommand check = new("SELECT COUNT(*) FROM Part_Entry WHERE Para_No=@Para_No AND ID!=@ID", con);
                check.Parameters.AddWithValue("@Para_No", paraNo);
                check.Parameters.AddWithValue("@ID", id);
                int count = (int)check.ExecuteScalar();
                if (count > 0)
                {
                    MessageBox.Show("Part Number already exists.", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                if (activePart)
                {
                    using SqlCommand reset = new("UPDATE Part_Entry SET ActivePart=0 WHERE ActivePart=1 AND ID!=@ID", con);
                    reset.Parameters.AddWithValue("@ID", id);
                    reset.ExecuteNonQuery();
                }

                using SqlCommand update = new("UPDATE Part_Entry SET Para_No=@Para_No, Para_Name=@Para_Name, ActivePart=@ActivePart, ID_Value=@ID_Value, BOT_Value=@BOT_Value WHERE ID=@ID", con);
                update.Parameters.AddWithValue("@Para_No", paraNo);
                update.Parameters.AddWithValue("@Para_Name", paraName);
                update.Parameters.AddWithValue("@ActivePart", activePart ? 1 : 0);
                update.Parameters.AddWithValue("@ID_Value", idValue);
                update.Parameters.AddWithValue("@BOT_Value", botValue);
                update.Parameters.AddWithValue("@ID", id);
                update.ExecuteNonQuery();

                MessageBox.Show("Record updated successfully.", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
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
            if (dataGrid.SelectedItem is not DataRowView row) return;

            try
            {
                int id = Convert.ToInt32(row["ID"]);
                string paraNo = row["Para_No"].ToString() ?? "";

                if (MessageBox.Show($"Delete {paraNo}?", "Confirm", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.No)
                    return;

                using SqlConnection con = new(connectionString);
                con.Open();
                using SqlCommand del = new("DELETE FROM Part_Entry WHERE ID=@ID", con);
                del.Parameters.AddWithValue("@ID", id);
                del.ExecuteNonQuery();

                MessageBox.Show("Record deleted successfully.", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
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
                using SqlConnection con = new(connectionString);
                con.Open();

                string query = @"
                        SELECT 
                        ID,
                        Para_No,
                        Para_Name,
                        CAST(ActivePart AS BIT) AS IsActive,
                        ID_Value,
                        BOT_Value
                    FROM Part_Entry
                    ORDER BY ID
                    ";

                SqlDataAdapter da = new(query, con);
                DataTable dt = new();
                da.Fill(dt);

                // ✅ FIXED: Force refresh for color update
                dataGrid.ItemsSource = null;
                dataGrid.ItemsSource = dt.DefaultView;
                dataGrid.Items.Refresh();

                if (dataGrid.Columns.Count > 0)
                    dataGrid.Columns[0].Visibility = Visibility.Collapsed;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }


        private void ClearInputs()
        {
            txtPartNumber.Text = "";
            txtPartName.Text = "";
            chkActivePart.IsChecked = false;
            cmbCategory.SelectedIndex = 0;
            cmbVendor.SelectedIndex = 0;
            btnUpdate.IsEnabled = false;
            btnDelete.IsEnabled = false;
        }

        private void dataGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (dataGrid.SelectedItem is DataRowView row)
            {
                txtPartNumber.Text = row["Para_No"].ToString();
                txtPartName.Text = row["Para_Name"].ToString();
                chkActivePart.IsChecked = Convert.ToBoolean(row["IsActive"]);
                cmbCategory.SelectedValue = Convert.ToInt32(row["ID_Value"]);
                cmbVendor.SelectedValue = Convert.ToInt32(row["BOT_Value"]);
                btnUpdate.IsEnabled = true;
                btnDelete.IsEnabled = true;
            }
            else
            {
                ClearInputs();
            }
        }

        public class ComboItem
        {
            public int Value { get; set; }
            public string Text { get; set; }
        }
    }
}
