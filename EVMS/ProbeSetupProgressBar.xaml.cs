using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Animation;

namespace EVMS
{
    public partial class ProbeSetupProgressBar : UserControl
    {
        public ProbeSetupProgressBar()
        {
            InitializeComponent();
        }

        // Title DependencyProperty
        public static readonly DependencyProperty TitleProperty =
            DependencyProperty.Register(
                nameof(Title),
                typeof(string),
                typeof(ProbeSetupProgressBar),
                new PropertyMetadata("Probe"));

        public string Title
        {
            get => (string)GetValue(TitleProperty);
            set => SetValue(TitleProperty, value);
        }

        // ID DependencyProperty
        public static readonly DependencyProperty IDProperty =
            DependencyProperty.Register(
                nameof(ID),
                typeof(string),
                typeof(ProbeSetupProgressBar),
                new PropertyMetadata("0"));

        public string ID
        {
            get => (string)GetValue(IDProperty);
            set => SetValue(IDProperty, value);
        }

        // Stroke DependencyProperty
        public static readonly DependencyProperty StrokeProperty =
            DependencyProperty.Register(
                nameof(Stroke),
                typeof(string),
                typeof(ProbeSetupProgressBar),
                new PropertyMetadata("0"));

        public string Stroke
        {
            get => (string)GetValue(StrokeProperty);
            set => SetValue(StrokeProperty, value);
        }

        // Value DependencyProperty
        public static readonly DependencyProperty ValueProperty =
            DependencyProperty.Register(
                nameof(Value),
                typeof(double),
                typeof(ProbeSetupProgressBar),
                new PropertyMetadata(0.0, OnValueChanged));

        public double Value
        {
            get => (double)GetValue(ValueProperty);
            set => SetValue(ValueProperty, value);
        }

        private static void OnValueChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is ProbeSetupProgressBar control && control.Bar != null)
            {
                double newValue = (double)e.NewValue;

                // Animate the ProgressBar value smoothly
                var animation = new DoubleAnimation
                {
                    From = control.Bar.Value,
                    To = newValue,
                    Duration = TimeSpan.FromMilliseconds(10),
                    EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseOut }
                };
                control.Bar.BeginAnimation(ProgressBar.ValueProperty, animation);

                // Color logic based on the new value thresholds
                if (newValue < (0.2 / 2.0) * 100 || newValue > (1.2 / 2.0) * 100)
                {
                    control.Bar.Foreground = new SolidColorBrush(Colors.Red);
                }
                else
                {
                    var gradientBrush = new LinearGradientBrush
                    {
                        StartPoint = new Point(0, 1),
                        EndPoint = new Point(0, 0)
                    };
                    gradientBrush.GradientStops.Add(new GradientStop(Color.FromRgb(0x4C, 0xAF, 0x50), 0)); // Green
                    gradientBrush.GradientStops.Add(new GradientStop(Color.FromRgb(0x21, 0x96, 0xF3), 1)); // Blue
                    control.Bar.Foreground = gradientBrush;
                }
            }
        }

        // Status DependencyProperty
        public static readonly DependencyProperty StatusProperty =
            DependencyProperty.Register(
                nameof(Status),
                typeof(string),
                typeof(ProbeSetupProgressBar),
                new PropertyMetadata(string.Empty));

        public string Status
        {
            get => (string)GetValue(StatusProperty);
            set => SetValue(StatusProperty, value);
        }
    }
}
