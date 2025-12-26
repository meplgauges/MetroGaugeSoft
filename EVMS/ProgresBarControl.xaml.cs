using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace EVMS
{
    public partial class ProgresBarControl : UserControl
    {
        public ProgresBarControl()
        {
            InitializeComponent();
            SizeChanged += ProgresBarControl_SizeChanged;
            Loaded += ProgresBarControl_Loaded;

        }
        private void ProgresBarControl_Loaded(object sender, RoutedEventArgs e)
        {
            // Initialize Value to Min or 0, depending on your range
            Value = Min;

            // Set fills to zero height initially
            AboveFill.Height = 0;
            BelowFill.Height = 0;

            BarValue.Text = "0.000";
        }

        private void ProgresBarControl_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            UpdateFill(Value);
        }

        public static readonly DependencyProperty MinProperty =
            DependencyProperty.Register(nameof(Min), typeof(double), typeof(ProgresBarControl), new PropertyMetadata(0.0));

        public static readonly DependencyProperty MeanProperty =
            DependencyProperty.Register(nameof(Mean), typeof(double), typeof(ProgresBarControl), new PropertyMetadata(50.0));

        public static readonly DependencyProperty MaxProperty =
            DependencyProperty.Register(nameof(Max), typeof(double), typeof(ProgresBarControl), new PropertyMetadata(100.0));

        public static readonly DependencyProperty ValueProperty =
            DependencyProperty.Register(nameof(Value), typeof(double), typeof(ProgresBarControl),
                new PropertyMetadata(0.0, OnValueChanged));

        public static readonly DependencyProperty TitleProperty =
            DependencyProperty.Register(nameof(Title), typeof(string), typeof(ProgresBarControl),
                new PropertyMetadata(string.Empty));

        public double Min
        {
            get => (double)GetValue(MinProperty);
            set => SetValue(MinProperty, value);
        }

        public double Mean
        {
            get => (double)GetValue(MeanProperty);
            set => SetValue(MeanProperty, value);
        }

        public double Max
        {
            get => (double)GetValue(MaxProperty);
            set => SetValue(MaxProperty, value);
        }

        public double Value
        {
            get => (double)GetValue(ValueProperty);
            set => SetValue(ValueProperty, value);
        }

        public string Title
        {
            get => (string)GetValue(TitleProperty);
            set => SetValue(TitleProperty, value);
        }

        private static void OnValueChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is ProgresBarControl control)
            {
                control.UpdateFill((double)e.NewValue);
            }
        }

        public void ResetVisuals()
        {
            AboveFill.Height = 0;
            AboveFill.Margin = new Thickness(0);
            BelowFill.Height = 0;
            BelowFill.Margin = new Thickness(0);

            // Optional: reset the fill color to a neutral or transparent color
            SolidColorBrush neutralBrush = new SolidColorBrush(Colors.Transparent);
            AboveFill.Fill = neutralBrush;
            BelowFill.Fill = neutralBrush;

            // Reset the text display
            BarValue.Text = "0.000";
        }


        private void UpdateFill(double value)
        {

            if (value == 0)
            {
                // Reset fills
                AboveFill.Height = 0;
                AboveFill.Margin = new Thickness(0);
                BelowFill.Height = 0;
                BelowFill.Margin = new Thickness(0);

                // Optional: reset color to transparent or some neutral color
                //var neutralBrush = new SolidColorBrush(Colors.Transparent); // Or choose another neutral
                //AboveFill.Fill = neutralBrush;
                //BelowFill.Fill = neutralBrush;

                // Set value display to zero with appropriate formatting
                BarValue.Text = "0.000";
                return;
            }
            // Validate range
            if (Max <= Min || Mean < Min || Mean > Max)
            {
                AboveFill.Height = 0;
                BelowFill.Height = 0;
                BarValue.Text = value.ToString("0.###");
                return;
            }

            // Check if value is out of range
            bool isOutOfRange = value < Min || value > Max;

            // Clamp value for drawing
            double clampedValue = Math.Max(Min, Math.Min(Max, value));
            BarValue.Text = value.ToString("F3");

            // ---- determine total height of the visual track ----
            double totalHeight = 150.0; // fallback (matches your XAML track height)

            if (AboveFill.Parent is FrameworkElement parent)
            {
                if (parent.ActualHeight > 0)
                    totalHeight = parent.ActualHeight;
                else if (parent is Panel panel)
                {
                    foreach (UIElement child in panel.Children)
                    {
                        if (child is Border b && b.ActualHeight > 0)
                        {
                            totalHeight = b.ActualHeight;
                            break;
                        }
                    }
                }
            }

            if (totalHeight <= 0) totalHeight = 150.0;
            double halfHeight = totalHeight / 2.0;

            // ---- compute fill fractions ----
            double fillAboveFraction = 0.0, fillBelowFraction = 0.0;

            if (clampedValue > Mean)
                fillAboveFraction = (clampedValue - Mean) / (Max - Mean);
            else if (clampedValue < Mean)
                fillBelowFraction = (Mean - clampedValue) / (Mean - Min);

            fillAboveFraction = Math.Clamp(fillAboveFraction, 0, 1);
            fillBelowFraction = Math.Clamp(fillBelowFraction, 0, 1);

            double pixelAboveHeight = fillAboveFraction * halfHeight;
            double pixelBelowHeight = fillBelowFraction * halfHeight;

            // ---- POSITIONING ----
            AboveFill.VerticalAlignment = VerticalAlignment.Bottom;
            AboveFill.Height = pixelAboveHeight;
            AboveFill.Margin = new Thickness(0, 0, 0, halfHeight);

            BelowFill.VerticalAlignment = VerticalAlignment.Top;
            BelowFill.Height = pixelBelowHeight;
            BelowFill.Margin = new Thickness(0, halfHeight, 0, 0);

            // ---- COLOR LOGIC ----
            SolidColorBrush fillColor = isOutOfRange
                ? new SolidColorBrush(Colors.Red)   // out of range
                : new SolidColorBrush(Colors.Green); // within range

            AboveFill.Fill = fillColor;
            BelowFill.Fill = fillColor;

            AboveFill.InvalidateMeasure();
            BelowFill.InvalidateMeasure();
        }
    }
}
