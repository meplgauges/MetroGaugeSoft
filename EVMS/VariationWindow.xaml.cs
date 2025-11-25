using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace EVMS
{
    public partial class VariationWindow : Window
    {
        public VariationWindow(string parameterName, List<double> values)
        {
            InitializeComponent();

            TitleText.Text = $"{parameterName}";

            if (values == null || values.Count == 0)
            {
                var noDataText = new TextBlock
                {
                    Text = "No valid readings found.",
                    FontSize = 18,
                    Foreground = System.Windows.Media.Brushes.Gray,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    Margin = new Thickness(0, 20, 0, 0)
                };
                ReadingsStack.Children.Add(noDataText);
                VariationText.Text = string.Empty;
                return;
            }

            // Add readings dynamically
            for (int i = 0; i < values.Count; i++)
            {
                var tb = new TextBlock
                {
                    Text = $"{i + 1}. {values[i]:F3}",
                    FontSize = 18,
                    Margin = new Thickness(0, 2, 0, 2),
                    Foreground = System.Windows.Media.Brushes.Black
                };
                ReadingsStack.Children.Add(tb);
            }

            double max = values.Max();
            double min = values.Min();
            double variation = max - min;

            VariationText.Text = $"Max: {max:F3}   Min: {min:F3}   Δ: {variation:F3}";
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

    }
}
