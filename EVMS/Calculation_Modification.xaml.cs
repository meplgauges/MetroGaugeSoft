using EVMS.Service;
using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;

namespace EVMS
{
    public partial class Calculation_Modification : UserControl
    {
        private DataStorageService? dataStorageService;
        private List<ParameterConfigItem> parameterItems = new List<ParameterConfigItem>();

        public Calculation_Modification()
        {
            InitializeComponent();
            this.Focusable = true;
            this.Focus();

            dataStorageService = new DataStorageService();

            this.Loaded += Calculation_Modification_Loaded;
            this.PreviewKeyDown += Calculation_Modification_PreviewKeyDown;
        }

        private void Calculation_Modification_Loaded(object? sender, RoutedEventArgs e)
        {
            this.Focusable = true;
            this.IsTabStop = true;
            Keyboard.Focus(this);
            FocusManager.SetFocusedElement(Window.GetWindow(this)!, this);

            // ✅ Auto-load first active part
            AutoLoadFirstActivePart();
        }

        private void Calculation_Modification_PreviewKeyDown(object sender, KeyEventArgs e)
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
                    var homePage = new Dashboard
                    {
                        HorizontalAlignment = HorizontalAlignment.Stretch,
                        VerticalAlignment = VerticalAlignment.Stretch
                    };
                    mainContentGrid.Children.Add(homePage);
                }
            }
        }

        private class ParameterConfigItem
        {
            public string? Para_No { get; set; }
            public string? Parameter { get; set; }
            public string? ShortName { get; set; }
            public int Sign_Change { get; set; }
            public double Compensation { get; set; }
        }

        // ✅ Auto-load first active part
        private void AutoLoadFirstActivePart()
        {
            try
            {
                var activeParts = dataStorageService!.GetActiveParts();

                if (activeParts.Count > 0)
                {
                    // Get first active part
                    string partNumber = activeParts[0].Para_No!;
                    //PartNumberTextBox.Text = partNumber;

                    // Load parameters for this part
                    LoadParametersFromService(partNumber);
                    GenerateParameterControls();
                }
                else
                {
                    MessageBox.Show("No active parts found.", "No Data",
                        MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading active parts: " + ex.Message, "Database Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }


        private void LoadParametersFromService(string partNumber)
        {
            parameterItems.Clear();
            try
            {
                var configList = dataStorageService!.GetPartConfigByPartNumber(partNumber);

                foreach (var config in configList)
                {
                    parameterItems.Add(new ParameterConfigItem
                    {
                        Para_No = config.Para_No,
                        Parameter = config.Parameter,
                        ShortName = config.ShortName,
                        Sign_Change = config.Sign_Change,
                        Compensation = config.Compensation
                    });
                }

                if (parameterItems.Count == 0)
                {
                    MessageBox.Show("No parameters found for the specified Part Number.", "No Data",
                        MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading parameters: " + ex.Message, "Database Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void GenerateParameterControls()
        {
            try
            {
                ToggleGrid1.Children.Clear();

                foreach (var item in parameterItems)
                {
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
                    rowGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(2, GridUnitType.Star) });
                    rowGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(3, GridUnitType.Star) });
                    rowGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(3, GridUnitType.Star) });
                    rowGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

                    // Column 1: Parameter Name
                    TextBlock paramText = new TextBlock
                    {
                        Text = item.Parameter,
                        Foreground = Brushes.Black,
                        FontWeight = FontWeights.Bold,
                        FontSize = 16,
                        VerticalAlignment = VerticalAlignment.Center,
                        HorizontalAlignment = HorizontalAlignment.Center, // ✅ Center horizontally
                        TextAlignment = TextAlignment.Center, // ✅ Center text alignment
                        Margin = new Thickness(10, 0, 0, 0)
                    };
                    Grid.SetColumn(paramText, 0);


                    // Column 2: Modern Toggle Switch with Sign
                    ToggleButton signToggle = new ToggleButton
                    {
                        Tag = item,
                        HorizontalAlignment = HorizontalAlignment.Center,
                        VerticalAlignment = VerticalAlignment.Center,
                        IsChecked = item.Sign_Change == 1,
                        Margin = new Thickness(5, 0, 5, 0),
                        Width = 40,
                        Height = 40,
                        Foreground = Brushes.White,
                        FontWeight = FontWeights.Bold,
                        FontSize = 24,
                        Content = item.Sign_Change == 1 ? "+" : "−",
                        BorderThickness = new Thickness(0),
                        Padding = new Thickness(0)
                    };

                    // ✅ Set initial background color
                    signToggle.Background = item.Sign_Change == 1 ? new SolidColorBrush(Color.FromRgb(76, 175, 80)) : new SolidColorBrush(Color.FromRgb(244, 67, 54));

                    // ✅ Apply rounded corner style
                    var toggleStyle = new Style(typeof(ToggleButton));
                    var controlTemplate = new ControlTemplate(typeof(ToggleButton));

                    var mainBorder = new FrameworkElementFactory(typeof(Border));
                    mainBorder.SetValue(Border.BackgroundProperty, new TemplateBindingExtension(ToggleButton.BackgroundProperty));
                    mainBorder.SetValue(Border.CornerRadiusProperty, new CornerRadius(8));
                    mainBorder.SetValue(Border.BorderThicknessProperty, new Thickness(0));

                    var contentPresenter = new FrameworkElementFactory(typeof(ContentPresenter));
                    contentPresenter.SetValue(ContentPresenter.HorizontalAlignmentProperty, HorizontalAlignment.Center);
                    contentPresenter.SetValue(ContentPresenter.VerticalAlignmentProperty, VerticalAlignment.Center);

                    mainBorder.AppendChild(contentPresenter);
                    controlTemplate.VisualTree = mainBorder;
                    toggleStyle.Setters.Add(new Setter(Control.TemplateProperty, controlTemplate));

                    signToggle.Style = toggleStyle;

                    signToggle.Checked += (s, e) =>
                    {
                        var btn = (ToggleButton)s;
                        btn.Background = new SolidColorBrush(Color.FromRgb(76, 175, 80)); // Green
                        btn.Content = "+";
                        UpdateSignChangeInService(item.Para_No!, item.Parameter!, 1);
                    };
                    signToggle.Unchecked += (s, e) =>
                    {
                        var btn = (ToggleButton)s;
                        btn.Background = new SolidColorBrush(Color.FromRgb(244, 67, 54)); // Red
                        btn.Content = "−";
                        UpdateSignChangeInService(item.Para_No!, item.Parameter!, 0);
                    };

                    Grid.SetColumn(signToggle, 1);


                    // Column 3: Compensation TextBox
                    TextBox compensationBox = new TextBox
                    {
                        Text = item.Compensation.ToString("F3"),
                        Tag = item,
                        VerticalAlignment = VerticalAlignment.Center,
                        HorizontalAlignment = HorizontalAlignment.Stretch,
                        Height = 35,
                        Margin = new Thickness(5, 0, 5, 0),
                        FontSize = 14,
                        FontWeight = FontWeights.Bold,
                        TextAlignment = TextAlignment.Center,
                        Width = 150
                    };

                    try
                    {
                        compensationBox.Style = (Style)FindResource("ModernFilledTextBox");
                    }
                    catch { }

                    Grid.SetColumn(compensationBox, 2);

                    // Column 4: Update Button
                    Button updateBtn = new Button
                    {
                        Content = "ADD",
                        Tag = new Tuple<ParameterConfigItem, TextBox>(item, compensationBox),
                        Height = 35,
                        Width=120,
                        Margin = new Thickness(5, 0, 5, 0),
                        FontWeight = FontWeights.Bold
                    };

                    try
                    {
                        updateBtn.Style = (Style)FindResource("GlossyBlueButton");
                    }
                    catch { }

                    updateBtn.Click += UpdateCompensationButton_Click;
                    Grid.SetColumn(updateBtn, 3);

                    rowGrid.Children.Add(paramText);
                    rowGrid.Children.Add(signToggle);
                    rowGrid.Children.Add(compensationBox);
                    rowGrid.Children.Add(updateBtn);

                    card.Child = rowGrid;
                    ToggleGrid1.Children.Add(card);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error generating controls: {ex.Message}", "UI Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }




        private void UpdateSignChangeInService(string paraNo, string Parameter, int signChangeValue)
        {
            try
            {
                bool success = dataStorageService!.UpdateSignChange(paraNo, Parameter, signChangeValue);
                if (!success)
                {
                    MessageBox.Show("Failed to update Sign_Change.", "Update Error",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error updating Sign_Change: " + ex.Message, "Database Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void UpdateCompensationButton_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is Tuple<ParameterConfigItem, TextBox> tuple)
            {
                var item = tuple.Item1;
                var textBox = tuple.Item2;

                if (!double.TryParse(textBox.Text.Trim(), out double compensationValue))
                {
                    MessageBox.Show("Please enter a valid numeric value for Compensation.", "Input Error",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                try
                {
                    bool success = dataStorageService!.UpdateCompensation(item.Para_No!, item.Parameter!, compensationValue);
                    if (success)
                    {
                        MessageBox.Show("Compensation updated successfully!", "Success",
                            MessageBoxButton.OK, MessageBoxImage.Information);
                        item.Compensation = compensationValue;
                    }
                    else
                    {
                        MessageBox.Show("Failed to update Compensation.", "Update Error",
                            MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error updating Compensation: " + ex.Message, "Database Error",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }
    }
}
