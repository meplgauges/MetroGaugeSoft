using ActUtlType64Lib;
using Microsoft.Data.SqlClient;
using System.Configuration;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;

namespace EVMS
{
    public partial class IO_Controle_page : UserControl
    {
        private IActUtlType64 plc = new ActUtlType64Class();
        private bool isConnected = false;
        private DispatcherTimer monitorTimer = new DispatcherTimer();
        private readonly string? connectionString;
        private Dictionary<string, Button> inputButtons = new Dictionary<string, Button>();
        private Dictionary<string, ToggleButton> outputButtons = new Dictionary<string, ToggleButton>();
        private List<IODevice> inputDevices = new List<IODevice>();
        private List<IODevice> outputDevices = new List<IODevice>();

        // Class representing an IO device
        public IO_Controle_page()
        {
            try
            {
                InitializeComponent();

                // Initialize connection string with error handling
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["EVMSDb"]?.ConnectionString;
                    if (string.IsNullOrEmpty(connectionString))
                    {
                        throw new InvalidOperationException("Database connection string 'EVMSDb' not found in configuration.");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Configuration error: {ex.Message}", "Configuration Error",
                        System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
                    return;
                }

                this.Loaded += SettingsPage_Loaded;

                // ✅ Register ESC key handler
                this.PreviewKeyDown += SettingsPage_PreviewKeyDown;

                // Initialize database
                InitializeDatabase();

                // Setup events with error handling
                SetupEventHandlers();

                // Setup timer with error handling
                SetupTimer();

                // Setup registration form buttons
                SetupRegistrationButtons();

                // Load devices from database and generate buttons
                LoadDevicesFromDatabase();
                GenerateButtons();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Initialization error: {ex.Message}", "Initialization Error",
                    System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
            }
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

        // Class representing an IO device
        public class IODevice
        {
            public int ID { get; set; }
            public string Description { get; set; }
            public string? Bit { get; set; }
            public bool IsInput { get; set; }
            public bool IsOutput { get; set; }
        }
        // Setup all UI event handlers
        private void SetupEventHandlers()
        {
            try
            {
                PowerToggle.Checked += PowerToggle_Checked;
                PowerToggle.Unchecked += PowerToggle_Unchecked;
                ColumnSelector.SelectionChanged += ColumnSelector_SelectionChanged;
                ColumnSelector1.SelectionChanged += ColumnSelector1_SelectionChanged;
                ColumnSelector.SelectedIndex = 1;
                ColumnSelector1.SelectedIndex = 1;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error setting up event handlers: {ex.Message}", "Setup Error",
                    System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
            }
        }

        // Setup the timer used to monitor IO states periodically
        private void SetupTimer()
        {
            try
            {
                monitorTimer.Interval = TimeSpan.FromMilliseconds(500);
                monitorTimer.Tick += MonitorTimer_Tick;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error setting up timer: {ex.Message}", "Timer Error",
                    System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
            }
        }

        // Initialize the IODevices table in the database if it does not exist
        private void InitializeDatabase()
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    string createTableQuery = @"
                        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='IODevices' AND xtype='U')
                        CREATE TABLE IODevices (
                            ID INT IDENTITY(1,1) PRIMARY KEY,
                            Description NVARCHAR(100),
                            Bit NVARCHAR(10),
                            IsInput BIT,
                            IsOutput BIT
                        )";
                    using (SqlCommand cmd = new SqlCommand(createTableQuery, conn))
                    {
                        cmd.ExecuteNonQuery();
                    }
                }
            }
            catch (SqlException sqlEx)
            {
                MessageBox.Show($"Database SQL error: {sqlEx.Message}\nError Number: {sqlEx.Number}",
                    "Database SQL Error", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
            }
            catch (InvalidOperationException ioEx)
            {
                MessageBox.Show($"Database connection error: {ioEx.Message}", "Connection Error",
                    System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Database initialization error: {ex.Message}", "Database Error",
                    System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
            }
        }

        // Setup event handlers for add, update, and delete device buttons in the registration form
        private void SetupRegistrationButtons()
        {
            try
            {
                var idTextBox = FindName("IdTextBox") as TextBox;
                // var descTextBox = FindName("DescTextBox") as TextBox;
                // var bitTextBox = FindName("BitTextBox") as TextBox;
                // var inputCheckBox = FindName("InputCheckBox") as CheckBox;
                // var outputCheckBox = FindName("OutputCheckBox") as CheckBox;
                var addButton = FindName("AddButton") as Button;
                var updateButton = FindName("UpdateButton") as Button;
                var deleteButton = FindName("DeleteButton") as Button;

                if (addButton != null)
                {
                    addButton.Click += (s, e) =>
                    {
                        string? description = DescTextBox?.Text;
                        string? bit = BitTextBox?.Text;
                        bool isInput = InputCheckBox?.IsChecked == true;
                        bool isOutput = OutputCheckBox?.IsChecked == true;
                        AddDevice(description, bit, isInput, isOutput);
                    };
                }

                if (updateButton != null)
                {
                    updateButton.Click += (s, e) =>
                    {
                        // Retrieve Id and other parameters for update
                        string? description = DescTextBox?.Text;
                        string? bit = BitTextBox?.Text;
                        bool isInput = InputCheckBox?.IsChecked == true;
                        bool isOutput = OutputCheckBox?.IsChecked == true;

                        //  UpdateDevice(description, bit, isInput, isOutput);
                    };
                }

                if (deleteButton != null)
                {
                    deleteButton.Click += (s, e) =>
                    {
                        // Delete based on description or Id as required
                        string? description = DescTextBox?.Text;
                        DeleteDevice(description);
                    };
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error setting up registration buttons: {ex.Message}", "Setup Error",
                    System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
            }
        }

        // Load all devices from the database into input and output device lists
        private void LoadDevicesFromDatabase()
        {
            try
            {
                inputDevices.Clear();
                outputDevices.Clear();

                using SqlConnection conn = new(connectionString);
                conn.Open();
                string query = "SELECT * FROM IODevices";

                using SqlCommand cmd = new(query, conn);
                using SqlDataReader reader = cmd.ExecuteReader();

                while (reader.Read())
                {
                    try
                    {
                        var device = new IODevice
                        {
                            ID = reader["ID"] != DBNull.Value ? (int)reader["ID"] : 0,
                            Description = reader["Description"]?.ToString() ?? "Unknown",
                            Bit = reader["Bit"]?.ToString() ?? "Unknown",
                            IsInput = reader["IsInput"] != DBNull.Value && (bool)reader["IsInput"],
                            IsOutput = reader["IsOutput"] != DBNull.Value && (bool)reader["IsOutput"]
                        };

                        if (device.IsInput)
                            inputDevices.Add(device);
                        if (device.IsOutput)
                            outputDevices.Add(device);
                    }
                    catch (Exception rowEx)
                    {
                        MessageBox.Show($"Error reading device row: {rowEx.Message}", "Data Error",
                            MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                }
            }
            catch (SqlException sqlEx)
            {
                MessageBox.Show($"Database SQL error while loading devices: {sqlEx.Message}", "Database Error",
                    System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
            }
            catch (InvalidOperationException ioEx)
            {
                MessageBox.Show($"Database connection error: {ioEx.Message}", "Connection Error",
                    System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading devices: {ex.Message}", "Database Error",
                    System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
            }
        }

        // Add a new device record to the database and refresh UI
        private void AddDevice(string description, string bit, bool isInput, bool isOutput)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(description) || string.IsNullOrWhiteSpace(bit))
                {
                    MessageBox.Show("Description and Bit are required fields.", "Validation Error",
                        System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                    return;
                }

                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    string query = "INSERT INTO IODevices (Description, Bit, IsInput, IsOutput) VALUES (@desc, @bit, @input, @output)";
                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    {
                        cmd.Parameters.AddWithValue("@desc", description);
                        cmd.Parameters.AddWithValue("@bit", bit);
                        cmd.Parameters.AddWithValue("@input", isInput);
                        cmd.Parameters.AddWithValue("@output", isOutput);
                        cmd.ExecuteNonQuery();
                    }
                }

                MessageBox.Show("Device added successfully!", "Success",
                    System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);

                LoadDevicesFromDatabase();
                GenerateButtons();
            }
            catch (SqlException sqlEx)
            {
                if (sqlEx.Number == 2627) // Unique constraint violation
                {
                    MessageBox.Show($"Device with bit '{bit}' already exists.", "Duplicate Error",
                        System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                }
                else
                {
                    MessageBox.Show($"Database error: {sqlEx.Message}", "Database Error",
                        System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error adding device: {ex.Message}", "Error",
                    System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
            }
        }


        //private void UpdateDevice( string description, string bit, bool isInput, bool isOutput)
        //{
        //    try
        //    {
        //        if (string.IsNullOrWhiteSpace(description) || string.IsNullOrWhiteSpace(bit))
        //        {
        //            MessageBox.Show("ID, Description, and Bit are required.", "Validation Error",
        //                MessageBoxButton.OK, MessageBoxImage.Warning);
        //            return;
        //        }
        //        using (SqlConnection conn = new SqlConnection(connectionString))
        //        {
        //            conn.Open();
        //            string query = "UPDATE IODevices SET Description = @desc, Bit = @bit, IsInput = @input, IsOutput = @output";
        //            using (SqlCommand cmd = new SqlCommand(query, conn))
        //            {
        //                cmd.Parameters.AddWithValue("@desc", description);
        //                cmd.Parameters.AddWithValue("@bit", bit);
        //                cmd.Parameters.AddWithValue("@input", isInput);
        //                cmd.Parameters.AddWithValue("@output", isOutput);
        //                int rows = cmd.ExecuteNonQuery();
        //                if (rows > 0)
        //                {
        //                    MessageBox.Show("Device updated successfully!", "Success",
        //                        MessageBoxButton.OK, MessageBoxImage.Information);
        //                }
        //                else
        //                {
        //                    MessageBox.Show("Device not found.", "Error",
        //                        MessageBoxButton.OK, MessageBoxImage.Warning);
        //                }
        //            }
        //        }
        //        LoadDevicesFromDatabase();
        //        GenerateButtons();
        //    }
        //    catch (SqlException sqlEx)
        //    {
        //        if (sqlEx.Number == 2627) // Unique constraint violation
        //        {
        //            MessageBox.Show($"Device with bit '{bit}' already exists.", "Duplicate Error",
        //                MessageBoxButton.OK, MessageBoxImage.Warning);
        //        }
        //        else
        //        {
        //            MessageBox.Show($"Database error: {sqlEx.Message}", "Database Error",
        //                MessageBoxButton.OK, MessageBoxImage.Error);
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show($"Error updating device: {ex.Message}", "Error",
        //            MessageBoxButton.OK, MessageBoxImage.Error);
        //    }
        //}

        private void DeleteDevice(string description)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(description))
                {
                    MessageBox.Show("Description is required.", "Validation Error",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                using SqlConnection conn = new(connectionString);
                conn.Open();
                string query = "DELETE FROM IODevices WHERE Description = @desc";

                using SqlCommand cmd = new(query, conn);
                cmd.Parameters.AddWithValue("@desc", description);
                int rows = cmd.ExecuteNonQuery();

                string message = rows > 0
                    ? "Device deleted successfully!"
                    : "Device not found.";

                MessageBox.Show(message, rows > 0 ? "Success" : "Not Found",
                    MessageBoxButton.OK, rows > 0 ? MessageBoxImage.Information : MessageBoxImage.Warning);

                LoadDevicesFromDatabase();
                GenerateButtons();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error deleting device: {ex.Message}", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }




        // Generate buttons for input devices dynamically
        private void GenerateButtons()
        {
            try
            {
                GenerateInputButtons();
                GenerateOutputButtons();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error generating buttons: {ex.Message}", "UI Error",
                    System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
            }
        }

        // Generate buttons representing input devices and add them to the UI
        private void GenerateInputButtons()
        {
            try
            {
                ToggleGrid1.Children.Clear();
                inputButtons.Clear();

                foreach (var device in inputDevices)
                {
                    try
                    {
                        Button inputButton = new Button
                        {
                            Width = 110,
                            Height = 80,
                            Margin = new Thickness(20),
                            Tag = device,
                            Background = Brushes.White,
                            Foreground = Brushes.Black,
                            Opacity = 0.7,
                            FontWeight = FontWeights.Bold,
                            Content = new TextBlock
                            {
                                Text = device.Description,
                                TextWrapping = TextWrapping.Wrap,
                                TextAlignment = TextAlignment.Center,
                                HorizontalAlignment = HorizontalAlignment.Center,
                                VerticalAlignment = VerticalAlignment.Center
                            }
                        };


                        try
                        {
                            inputButton.Style = (Style)FindResource("ModernButtonStyle");
                        }
                        catch (ResourceReferenceKeyNotFoundException)
                        {
                            // Style not found, continue with default
                        }

                        inputButtons[device.Bit] = inputButton;
                        ToggleGrid1.Children.Add(inputButton);
                    }
                    catch (Exception buttonEx)
                    {
                        MessageBox.Show($"Error creating input button for {device.Bit}: {buttonEx.Message}",
                            "Button Creation Error", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error generating input buttons: {ex.Message}", "UI Error",
                    System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
            }
        }

        // Generate toggle buttons representing output devices and add them to the UI
        private void GenerateOutputButtons()
        {
            try
            {
                ToggleGrid.Children.Clear();
                outputButtons.Clear();

                foreach (var device in outputDevices)
                {
                    try
                    {
                        StackPanel container = new StackPanel
                        {
                            Orientation = Orientation.Vertical,
                            Margin = new Thickness(5),
                            HorizontalAlignment = HorizontalAlignment.Center,
                            VerticalAlignment = VerticalAlignment.Center
                        };

                        ToggleButton outputButton = new ToggleButton
                        {
                            Tag = device,
                            IsEnabled = false,
                            HorizontalAlignment = HorizontalAlignment.Center,
                            Margin = new Thickness(0, 0, 0, 5)
                        };

                        try
                        {
                            outputButton.Style = (Style)FindResource("ModernSwitchToggleStyle");
                        }
                        catch (ResourceReferenceKeyNotFoundException)
                        {
                            // Style not found, set default properties
                            outputButton.Width = 80;
                            outputButton.Height = 36;
                        }

                        TextBlock deviceLabel = new TextBlock
                        {
                            Text = $"{device.Description}",
                            Foreground = Brushes.White,
                            FontWeight = FontWeights.Bold,
                            FontSize = 12,
                            TextAlignment = TextAlignment.Center,
                            HorizontalAlignment = HorizontalAlignment.Center,
                            VerticalAlignment = VerticalAlignment.Center,
                            Margin = new Thickness(0, 10, 0, 0)
                        };

                        outputButton.Checked += OutputButton_Checked;
                        outputButton.Unchecked += OutputButton_Unchecked;
                        outputButtons[device.Bit] = outputButton;

                        container.Children.Add(outputButton);
                        container.Children.Add(deviceLabel);
                        ToggleGrid.Children.Add(container);
                    }
                    catch (Exception buttonEx)
                    {
                        MessageBox.Show($"Error creating output button for {device.Bit}: {buttonEx.Message}",
                            "Button Creation Error", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error generating output buttons: {ex.Message}", "UI Error",
                    System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
            }
        }

        // Event handler for changing the number of columns in output buttons grid
        private void ColumnSelector_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                if (ColumnSelector.SelectedItem is ComboBoxItem selectedItem)
                {
                    if (int.TryParse(selectedItem.Content.ToString(), out int columns))
                    {
                        ToggleGrid.Columns = columns;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error changing column selection: {ex.Message}", "UI Error",
                    System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
            }
        }

        // Event handler for changing the number of columns in input buttons grid
        private void ColumnSelector1_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                if (ColumnSelector1.SelectedItem is ComboBoxItem selectedItem)
                {
                    if (int.TryParse(selectedItem.Content.ToString(), out int columns))
                    {
                        ToggleGrid1.Columns = columns;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error changing column selection: {ex.Message}", "UI Error",
                    System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
            }
        }

        // Event handler for PLC connection toggle On (connects to PLC)
        private async void PowerToggle_Checked(object sender, RoutedEventArgs e)
        {
            PowerToggle.IsEnabled = false; // Prevent multiple clicks immediately

            try
            {
                plc.ActLogicalStationNumber = 1;

                var openTask = Task.Run(() =>
                {
                    try
                    {
                        return plc.Open();  // Blocking call
                    }
                    catch
                    {
                        return -1; // Exception indicator
                    }
                });

                // Wait for either openTask or timeout (3 seconds)
                var completedTask = await Task.WhenAny(openTask, Task.Delay(1000));

                if (completedTask == openTask)
                {
                    // Connection attempt finished within timeout
                    int result = openTask.Result;

                    if (result == 0)
                    {
                        isConnected = true;
                        MessageBox.Show("✅ Connected to PLC", "PLC Connection",
                            MessageBoxButton.OK, MessageBoxImage.Information);
                        PlcStatusText.Text = "PLC Connected ✅";
                        foreach (var button in outputButtons.Values)
                            button.IsEnabled = true;
                        monitorTimer.Start();
                        InitializeButtonStates();
                    }
                    else if (result == -1)
                    {
                        isConnected = false;
                        MessageBox.Show("❌ Failed to connect to PLC (exception). Please check the physical connection.",
                            "PLC Connection Error", MessageBoxButton.OK, MessageBoxImage.Error);
                        PlcStatusText.Text = "Connection Failed ❌";
                        PowerToggle.IsChecked = false;
                    }
                    else
                    {
                        isConnected = false;
                        MessageBox.Show($"❌ Failed to connect to PLC. Error code: {result}",
                            "PLC Connection Error", MessageBoxButton.OK, MessageBoxImage.Error);
                        PlcStatusText.Text = "Connection Failed ❌";
                        PowerToggle.IsChecked = false;
                    }
                }
                else
                {
                    // Timeout occurred first
                    isConnected = false;
                    MessageBox.Show("❌ PLC connection timed out after 3 seconds. Please check physical connection.",
                        "PLC Connection Timeout", MessageBoxButton.OK, MessageBoxImage.Error);
                    PlcStatusText.Text = "Connection Timeout ❌";
                    PowerToggle.IsChecked = false;
                }
            }
            catch (Exception ex)
            {
                isConnected = false;
                PowerToggle.IsChecked = false;
                MessageBox.Show($"PLC connection error: {ex.Message}", "PLC Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                PlcStatusText.Text = "Connection Error ❌";
            }
            finally
            {
                PowerToggle.IsEnabled = true;
            }
        }



        // Event handler for PLC connection toggle Off (disconnects from PLC)
        private void PowerToggle_Unchecked(object sender, RoutedEventArgs e)
        {
            try
            {
                if (isConnected)
                {
                    try
                    {
                        plc.Close();
                    }
                    catch (Exception closeEx)
                    {
                        MessageBox.Show($"Error closing PLC connection: {closeEx.Message}", "PLC Error",
                            System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                    }

                    isConnected = false;
                    MessageBox.Show("PLC Disconnected", "PLC Connection",
                        System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);
                    PlcStatusText.Text = "PLC Disconnected ❌";

                    monitorTimer.Stop();

                    foreach (var button in outputButtons.Values)
                    {
                        button.IsEnabled = false;
                        button.Background = Brushes.Gray;
                    }

                    foreach (var button in inputButtons.Values)
                    {
                        button.Background = Brushes.Gray;
                    }
                }
                else
                {
                    MessageBox.Show("PLC already disconnected.", "PLC Connection",
                        System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error during PLC disconnection: {ex.Message}", "PLC Error",
                    System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
            }
        }

        // Event handler when an output toggle button is checked (ON)
        private void OutputButton_Checked(object sender, RoutedEventArgs e)
        {
            try
            {
                if (sender is ToggleButton button && button.Tag is IODevice device)
                {
                    if (isConnected)
                    {
                        try
                        {
                            int result = plc.SetDevice(device.Bit, 1);
                            if (result == 0)
                            {
                                button.Background = Brushes.Green;
                            }
                            else
                            {
                                string errorMessage = GetPLCErrorMessage(result);
                                MessageBox.Show($"Failed to turn {device.Bit} ON.\nError code: {result}\nDescription: {errorMessage}",
                                    $"{device.Bit} Output Error", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
                                button.IsChecked = false;
                                button.Background = Brushes.Red;
                            }
                        }
                        catch (Exception plcEx)
                        {
                            MessageBox.Show($"PLC communication error for {device.Bit}: {plcEx.Message}", "PLC Error",
                                System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
                            button.IsChecked = false;
                        }
                    }
                    else
                    {
                        MessageBox.Show("PLC not connected.", "Connection Error",
                            System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                        button.IsChecked = false;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error in output button checked event: {ex.Message}", "UI Error",
                    System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
            }
        }


        // Event handler when an output toggle button is unchecked (OFF)
        private void OutputButton_Unchecked(object sender, RoutedEventArgs e)
        {
            try
            {
                if (sender is ToggleButton button && button.Tag is IODevice device)
                {
                    if (isConnected)
                    {
                        try
                        {
                            int result = plc.SetDevice(device.Bit, 0);
                            if (result == 0)
                            {
                                button.Background = Brushes.Red;
                            }
                            else
                            {
                                MessageBox.Show($"Failed to turn {device.Bit} OFF. Error code: {result}",
                                    $"{device.Bit} Output Error", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
                                button.IsChecked = true;
                                button.Background = Brushes.Green;
                            }
                        }
                        catch (Exception plcEx)
                        {
                            MessageBox.Show($"PLC communication error: {plcEx.Message}", "PLC Error",
                                System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
                            button.IsChecked = true;
                        }
                    }
                    else
                    {
                        MessageBox.Show("PLC not connected.", "Connection Error",
                            System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                        button.IsChecked = false;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error in output button unchecked event: {ex.Message}", "UI Error",
                    System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
            }
        }

        // Initialize output button states to reflect current PLC device statuses
        private void InitializeButtonStates()
        {
            try
            {
                foreach (var device in outputDevices)
                {
                    try
                    {
                        int value;
                        int result = plc.GetDevice(device.Bit, out value);
                        if (result == 0 && outputButtons.ContainsKey(device.Bit))
                        {
                            outputButtons[device.Bit].IsChecked = (value == 1);
                            outputButtons[device.Bit].Background = (value == 1) ? Brushes.Green : Brushes.Red;
                        }
                    }
                    catch (Exception deviceEx)
                    {
                        MessageBox.Show($"Error initializing state for {device.Bit}: {deviceEx.Message}",
                            "PLC Error", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error initializing button states: {ex.Message}", "Initialization Error",
                    System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
            }
        }

        // Returns the user-friendly error message based on PLC error code
        private string GetPLCErrorMessage(int errorCode)
        {
            switch (errorCode)
            {
                case 25174017: // 0x1800001
                    return "Device address out of range or device not available. Check if Y8 exists on your PLC.";
                case 25174018: // 0x1800002
                    return "Invalid data type specified.";
                case 25174019: // 0x1800003
                    return "Device address format error.";
                case 25174020: // 0x1800004
                    return "Data range error.";
                case 25174016: // 0x1800000
                    return "Communication timeout. Check PLC connection.";
                case 25169920: // 0x1801000
                    return "PLC is not in RUN mode. Set PLC to RUN mode.";
                case 25165824: // 0x1800000
                    return "Communication error. Check cable and settings.";
                default:
                    return $"Unknown PLC error. Check PLC documentation for error code: {errorCode}";
            }
        }

        // Timer tick event to periodically update input device button states from PLC
        private void MonitorTimer_Tick(object sender, EventArgs e)
        {
            try
            {
                if (isConnected)
                {
                    foreach (var device in inputDevices)
                    {
                        try
                        {
                            if (inputButtons.ContainsKey(device.Bit))
                            {
                                int value;
                                int result = plc.GetDevice(device.Bit, out value);
                                if (result == 0)
                                {
                                    inputButtons[device.Bit].Background = (value == 1) ? Brushes.Green : Brushes.Red;
                                }
                                else
                                {
                                    inputButtons[device.Bit].Background = Brushes.Gray;
                                }
                            }
                        }
                        catch (Exception)
                        {
                            // Log device-specific errors without showing popup (too frequent)
                            if (inputButtons.ContainsKey(device.Bit))
                            {
                                inputButtons[device.Bit].Background = Brushes.Gray;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Stop timer on critical error
                monitorTimer.Stop();
                MessageBox.Show($"Monitor timer error: {ex.Message}", "Monitor Error",
                    System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
            }
        }
    }
}
