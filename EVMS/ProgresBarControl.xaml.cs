using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace EVMS
{
    public partial class ProgresBarControl : UserControl
    {
        private bool _isInitialized;
        private bool _showMin;

        public ProgresBarControl()
        {
            InitializeComponent();
            Loaded += ProgresBarControl_Loaded;
            SizeChanged += ProgresBarControl_SizeChanged;
        }

        private void ProgresBarControl_Loaded(object sender, RoutedEventArgs e)
        {
            // Decide visibility ONCE
            DecideInitialMinVisibility();
            ApplyMinVisibility();

            // Initial visuals
            Value = Min;
            AboveFill.Height = 0;
            BelowFill.Height = 0;
            BarValue.Text = "0.000";
            MinValue.Text = "0.000";


            _isInitialized = true; // 🔒 LOCK VISIBILITY
        }

        private void ProgresBarControl_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            UpdateFill(Value);
        }

        // ---------------- DEPENDENCY PROPERTIES ----------------

        public static readonly DependencyProperty MinProperty =
            DependencyProperty.Register(
                nameof(Min),
                typeof(double),
                typeof(ProgresBarControl),
                new PropertyMetadata(0.0));

        public static readonly DependencyProperty MaxProperty =
            DependencyProperty.Register(
                nameof(Max),
                typeof(double),
                typeof(ProgresBarControl),
                new PropertyMetadata(100.0));

        public static readonly DependencyProperty MeanProperty =
            DependencyProperty.Register(
                nameof(Mean),
                typeof(double),
                typeof(ProgresBarControl),
                new PropertyMetadata(50.0));

        public static readonly DependencyProperty MinValueProperty =
            DependencyProperty.Register(
                nameof(MValue),
                typeof(double),
                typeof(ProgresBarControl),
                new PropertyMetadata(0.0, OnValueChanged));

        public static readonly DependencyProperty ValueProperty =
            DependencyProperty.Register(
                nameof(Value),
                typeof(double),
                typeof(ProgresBarControl),
                new PropertyMetadata(0.0, OnValueChanged));

        public static readonly DependencyProperty TitleProperty =
            DependencyProperty.Register(
                nameof(Title),
                typeof(string),
                typeof(ProgresBarControl),
                new PropertyMetadata(string.Empty));

        // ---------------- PROPERTIES ----------------

        public double Min
        {
            get => (double)GetValue(MinProperty);
            set => SetValue(MinProperty, value);
        }

        public double Max
        {
            get => (double)GetValue(MaxProperty);
            set => SetValue(MaxProperty, value);
        }

        public double Mean
        {
            get => (double)GetValue(MeanProperty);
            set => SetValue(MeanProperty, value);
        }

        public double MValue
        {
            get => (double)GetValue(MinValueProperty);
            set => SetValue(MinValueProperty, value);
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

        // ---------------- VISIBILITY LOGIC ----------------

        private void DecideInitialMinVisibility()
        {
            // Rule 1: Min == Max → SHOW
            if (Min == Max)
            {
                _showMin = true;
                return;
            }

            // Rule 2: Title-based
            if (string.IsNullOrWhiteSpace(Title))
            {
                _showMin = true;
                return;
            }

            string title = Title.Trim().ToUpperInvariant();

            _showMin = !(
                title.StartsWith("RN") ||
                title.StartsWith("TL") ||
                title.StartsWith("STEP") ||
                title.Contains("RUNOUT")
            );
        }

        private void ApplyMinVisibility()
        {
            var visibility = _showMin
                ? Visibility.Visible
                : Visibility.Collapsed;

            MinLabel.Visibility = visibility;
            MinValue.Visibility = visibility;
        }

        // ---------------- VALUE / FILL LOGIC ----------------

        private static void OnValueChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is ProgresBarControl c)
                c.UpdateFill((double)e.NewValue);
        }

        private void UpdateFill(double value)
        {
            // ----- TEXT -----
            BarValue.Text = value.ToString("F3");
            MinValue.Text = MValue.ToString("F3");

            // ----- ALWAYS RESET VISUALS FIRST -----
            AboveFill.Height = 0;
            BelowFill.Height = 0;
            AboveFill.Margin = new Thickness(0);
            BelowFill.Margin = new Thickness(0);

            var transparent = new SolidColorBrush(Colors.Transparent);
            AboveFill.Fill = transparent;
            BelowFill.Fill = transparent;

            // 🔥 1. RESET STATE (VERY IMPORTANT)
            if (value == 0)
                return;

            // 🔥 2. INVALID RANGE
            if (Max <= Min || Mean < Min || Mean > Max)
                return;

            // 🔥 3. OUT OF RANGE → RED, NO FILL
            // 🔥 3. OUT OF RANGE → RED, FULL BAR
            // 🔥 3. OUT OF RANGE → RED (DIRECTION AWARE)
            if (value < Min || value > Max)
            {
                double rTotalHeight = 150.0;
                if (AboveFill.Parent is FrameworkElement rFe && rFe.ActualHeight > 0)
                    rTotalHeight = rFe.ActualHeight;

                double rHalfHeight = rTotalHeight / 2.0;

                if (value > Max)
                {
                    // 🔴 ABOVE
                    AboveFill.Height = rHalfHeight;
                    AboveFill.VerticalAlignment = VerticalAlignment.Bottom;
                    AboveFill.Margin = new Thickness(0, 0, 0, rHalfHeight);
                    AboveFill.Fill = new SolidColorBrush(Colors.Red);
                }
                else
                {
                    // 🔴 BELOW
                    BelowFill.Height = rHalfHeight;
                    BelowFill.VerticalAlignment = VerticalAlignment.Top;
                    BelowFill.Margin = new Thickness(0, rHalfHeight, 0, 0);
                    BelowFill.Fill = new SolidColorBrush(Colors.Red);
                }

                return;
            }


            // 🔥 4. ZERO TOLERANCE (Min == Max) → SHOW MIDPOINT MARKER
            if (Min == Max)
            {
                double zTotalHeight = 150.0;
                if (AboveFill.Parent is FrameworkElement zFe && zFe.ActualHeight > 0)
                    zTotalHeight = zFe.ActualHeight;

                double zHalfHeight = zTotalHeight / 2.0;
                double markerHeight = 8;

                AboveFill.Height = markerHeight;
                AboveFill.VerticalAlignment = VerticalAlignment.Bottom;
                AboveFill.Margin = new Thickness(0, 0, 0, zHalfHeight - markerHeight / 2);
                AboveFill.Fill = new SolidColorBrush(Colors.Green);

                return;
            }


            // 🔥 4. IN RANGE → GREEN FILL
            double totalHeight = 150.0;
            if (AboveFill.Parent is FrameworkElement fe && fe.ActualHeight > 0)
                totalHeight = fe.ActualHeight;

            double halfHeight = totalHeight / 2.0;

            double above = 0, below = 0;

            if (value > Mean)
                above = (value - Mean) / (Max - Mean);
            else if (value < Mean)
                below = (Mean - value) / (Mean - Min);

            AboveFill.Height = Math.Clamp(above, 0, 1) * halfHeight;
            BelowFill.Height = Math.Clamp(below, 0, 1) * halfHeight;

            AboveFill.VerticalAlignment = VerticalAlignment.Bottom;
            BelowFill.VerticalAlignment = VerticalAlignment.Top;

            AboveFill.Margin = new Thickness(0, 0, 0, halfHeight);
            BelowFill.Margin = new Thickness(0, halfHeight, 0, 0);

            var green = new SolidColorBrush(Colors.Green);
            AboveFill.Fill = green;
            BelowFill.Fill = green;
        }

    }
}
