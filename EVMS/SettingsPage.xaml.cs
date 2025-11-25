using EVMS.Service;
using Microsoft.Data.SqlClient;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;

namespace EVMS
{
    public partial class SettingsPage : UserControl
    {
        private string? connectionString;
        private readonly DataStorageService _dataStorageService;

        private Dictionary<int, ToggleButton> outputButtons = new Dictionary<int, ToggleButton>();
        private List<ControlItem> outputDevices = new List<ControlItem>();

        public SettingsPage()
        {
            InitializeComponent();
            _dataStorageService = new DataStorageService();

            this.Focusable = true;
            this.Focus();
            connectionString = ConfigurationManager.ConnectionStrings["EVMSDb"]?.ConnectionString;
            if (string.IsNullOrEmpty(connectionString))
            {
                MessageBox.Show("Database connection string 'EVMSDb' not found.", "Config Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            AddButton.Click += AddButton_Click;
            DeleteButton.Click += DeleteButton_Click;
            UpdateButton.Click += UpdateButton_Click;

            this.Loaded += SettingsPage_Loaded;

            // ✅ Register ESC key handler
            this.PreviewKeyDown += SettingsPage_PreviewKeyDown;
            LoadAndGenerate();
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

            try
            {
                int count = _dataStorageService.GetReadingCount();
                ReadingCountTextBox.Text = count.ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading Reading Count: {ex.Message}");
            }
        }
        // ✅ Handles ESC key press to go back to HomePage

        private void UpdateButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (int.TryParse(ReadingCountTextBox.Text, out int newCount))
                {
                    _dataStorageService.UpdateReadingCount(newCount);
                    MessageBox.Show("Reading Count updated successfully!", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    MessageBox.Show("Please enter a valid number.", "Invalid Input", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error updating Reading Count: {ex.Message}");
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

                    var resultPage = new Dashboard
                    {
                        HorizontalAlignment = HorizontalAlignment.Stretch,
                        VerticalAlignment = VerticalAlignment.Stretch
                    };

                    mainContentGrid.Children.Add(resultPage);
                }
            }
        }


        private class ControlItem
        {
            public string? Description { get; set; }
            public int Bit { get; set; }
            public string? Code { get; set; }
        }

        private void LoadAndGenerate()
        {
            outputDevices = LoadItemsFromDatabase();
            GenerateInputButtons();
        }

        private List<ControlItem> LoadItemsFromDatabase()
        {
            var list = new List<ControlItem>();
            try
            {
                using (var conn = new SqlConnection(connectionString))
                {
                    string sql = "SELECT Description, Bit, Code FROM Controls";
                    var cmd = new SqlCommand(sql, conn);
                    conn.Open();
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            bool bitValue = reader.GetBoolean(1); // ✅ read BIT as bool
                            list.Add(new ControlItem
                            {
                                Description = reader.GetString(0),
                                Bit = bitValue ? 1 : 0 , // ✅ convert bool → int
                                Code = reader.GetString(2)
                            });
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading data: " + ex.Message, "Database Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
            return list;
        }


        private void GenerateInputButtons()
        {
            try
            {
                ToggleGrid1.Children.Clear();
                outputButtons.Clear();

                foreach (var device in outputDevices)
                {
                    // Each row container
                    Border card = new Border
                    {
                        Background = new SolidColorBrush(Color.FromArgb(60, 255, 255, 255)),
                        CornerRadius = new CornerRadius(10),
                        Margin = new Thickness(0, 6, 0, 6),
                        Padding = new Thickness(15, 10, 15, 10),
                        BorderBrush = Brushes.Gray,
                        BorderThickness = new Thickness(1)
                    };

                    Grid rowGrid = new Grid();
                    // Add three columns: Description, Code, Toggle
                    rowGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(3, GridUnitType.Star) });
                    rowGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(2, GridUnitType.Star) });
                    rowGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

                    // Description TextBlock - Left aligned
                    TextBlock descText = new TextBlock
                    {
                        Text = device.Description,
                        Foreground = Brushes.Black,
                        FontWeight = FontWeights.Bold,
                        FontSize = 14,
                        VerticalAlignment = VerticalAlignment.Center,
                        Margin = new Thickness(10, 0, 0, 0)
                    };
                    Grid.SetColumn(descText, 0);

                    // Code TextBlock - Center aligned
                    TextBlock codeText = new TextBlock
                    {
                        Text = device.Code,
                        Foreground = Brushes.White,
                        FontWeight = FontWeights.Normal,
                        FontSize = 14,
                        VerticalAlignment = VerticalAlignment.Center,
                        HorizontalAlignment = HorizontalAlignment.Left,
                    };
                    Grid.SetColumn(codeText, 1);

                    ToggleButton toggle = new ToggleButton
                    {
                        Tag = device,
                        HorizontalAlignment = HorizontalAlignment.Center,
                        VerticalAlignment = VerticalAlignment.Center,
                        IsChecked = device.Bit == 1,
                        Margin = new Thickness(0, 0, 10, 0)
                    };

                    try
                    {
                        toggle.Style = (Style)FindResource("ModernSwitchToggleStyle");
                    }
                    catch
                    {
                        toggle.Width = 60;
                        toggle.Height = 30;
                        toggle.Content = device.Bit == 1 ? "ON" : "OFF";
                    }

                    toggle.Checked += (s, e) => UpdateBitInDatabase(device.Description!, 1);
                    toggle.Unchecked += (s, e) => UpdateBitInDatabase(device.Description!, 0);

                    Grid.SetColumn(toggle, 2);

                    rowGrid.Children.Add(descText);
                    rowGrid.Children.Add(codeText);
                    rowGrid.Children.Add(toggle);

                    card.Child = rowGrid;
                    ToggleGrid1.Children.Add(card);

                    outputButtons[device.Bit] = toggle;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error generating output buttons: {ex.Message}", "UI Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }



        private void OutputButton_Checked(object sender, RoutedEventArgs e)
        {
            if (sender is ToggleButton toggle && toggle.Tag is ControlItem device)
            {
                UpdateBitInDatabase(device.Description!, 1);
            }
        }

        private void OutputButton_Unchecked(object sender, RoutedEventArgs e)
        {
            if (sender is ToggleButton toggle && toggle.Tag is ControlItem device)
            {
                UpdateBitInDatabase(device.Description!, 0);
            }
        }

        private void UpdateBitInDatabase(string description, int bitValue)
        {
            try
            {
                using (var conn = new SqlConnection(connectionString))
                {
                    string sql = "UPDATE Controls SET Bit = @bit WHERE Description = @desc";
                    var cmd = new SqlCommand(sql, conn);
                    cmd.Parameters.AddWithValue("@bit", bitValue);
                    cmd.Parameters.AddWithValue("@desc", description);
                    conn.Open();
                    cmd.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error updating Bit value: " + ex.Message, "Database Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void AddButton_Click(object sender, RoutedEventArgs e)
        {
            string description = DescTextBox.Text.Trim();
            string code = BitCode.Text.Trim(); // Get Code value
            if (string.IsNullOrEmpty(description))
            {
                MessageBox.Show("Please enter a description.", "Input Error",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (!int.TryParse(BitTextBox.Text.Trim(), out int bitInt) || (bitInt != 0 && bitInt != 1))
            {
                MessageBox.Show("Bit must be 0 or 1.", "Input Error",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (string.IsNullOrEmpty(code))
            {
                MessageBox.Show("Please enter a code.", "Input Error",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                using (var conn = new SqlConnection(connectionString))
                {
                    string sql = "INSERT INTO Controls (Description, Bit, Code) VALUES (@desc, @bit, @code)";
                    var cmd = new SqlCommand(sql, conn);
                    cmd.Parameters.AddWithValue("@desc", description);
                    cmd.Parameters.AddWithValue("@bit", bitInt);
                    cmd.Parameters.AddWithValue("@code", code);  // Add Code parameter

                    conn.Open();
                    cmd.ExecuteNonQuery();
                }
                LoadAndGenerate();
                ClearInputs();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error adding record: " + ex.Message, "Database Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        //private void UpdateButton_Click(object sender, RoutedEventArgs e)
        //{
        //    string description = DescTextBox.Text.Trim();
        //    if (string.IsNullOrEmpty(description))
        //    {
        //        MessageBox.Show("Please enter a description.", "Input Error",
        //            MessageBoxButton.OK, MessageBoxImage.Warning);
        //        return;
        //    }

        //    if (!int.TryParse(BitTextBox.Text.Trim(), out int bitInt) || (bitInt != 0 && bitInt != 1))
        //    {
        //        MessageBox.Show("Bit must be 0 or 1.", "Input Error",
        //            MessageBoxButton.OK, MessageBoxImage.Warning);
        //        return;
        //    }

        //    try
        //    {
        //        using (var conn = new SqlConnection(connectionString))
        //        {
        //            string sql = "UPDATE Controls SET Bit = @bit WHERE Description = @desc";
        //            var cmd = new SqlCommand(sql, conn);
        //            cmd.Parameters.AddWithValue("@desc", description);
        //            cmd.Parameters.AddWithValue("@bit", bitInt);
        //            conn.Open();
        //            int rows = cmd.ExecuteNonQuery();
        //            if (rows == 0)
        //            {
        //                MessageBox.Show("No record found to update.", "Update Error",
        //                    MessageBoxButton.OK, MessageBoxImage.Warning);
        //                return;
        //            }
        //        }
        //        LoadAndGenerate();
        //        ClearInputs();
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show("Error updating record: " + ex.Message, "Database Error",
        //            MessageBoxButton.OK, MessageBoxImage.Error);
        //    }
        //}

        private void DeleteButton_Click(object sender, RoutedEventArgs e)
        {
            string description = DescTextBox.Text.Trim();
            if (string.IsNullOrEmpty(description))
            {
                MessageBox.Show("Please enter a description to delete.", "Input Error",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                using (var conn = new SqlConnection(connectionString))
                {
                    string sql = "DELETE FROM Controls WHERE Description = @desc";
                    var cmd = new SqlCommand(sql, conn);
                    cmd.Parameters.AddWithValue("@desc", description);
                    conn.Open();
                    int rows = cmd.ExecuteNonQuery();
                    if (rows == 0)
                    {
                        MessageBox.Show("No record found to delete.", "Delete Error",
                            MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }
                }
                LoadAndGenerate();
                ClearInputs();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error deleting record: " + ex.Message, "Database Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ClearInputs()
        {
            DescTextBox.Text = "";
            BitTextBox.Text = "";
            BitCode.Text = "";
        }
    }
}
