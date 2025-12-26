using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Animation;
using System.Windows.Shapes;

namespace EVMS
{
    public partial class ProbeSetupProgressBar : UserControl
    {
        public ProbeSetupProgressBar()
        {
            InitializeComponent();
            UpdateMidLine(); // Ensure mid line is positioned correctly at start
        }

        public override void OnApplyTemplate()
        {
            base.OnApplyTemplate();
            UpdateMidLine();
        }

        // ---------------------------------------------------
        // Dependency Properties
        // ---------------------------------------------------

        #region Title
        public static readonly DependencyProperty TitleProperty =
            DependencyProperty.Register(nameof(Title), typeof(string),
            typeof(ProbeSetupProgressBar), new PropertyMetadata("Probe"));

        public string Title
        {
            get => (string)GetValue(TitleProperty);
            set => SetValue(TitleProperty, value);
        }
        #endregion

        #region ID
        public static readonly DependencyProperty IDProperty =
            DependencyProperty.Register(nameof(ID), typeof(string),
            typeof(ProbeSetupProgressBar), new PropertyMetadata("0"));

        public string ID
        {
            get => (string)GetValue(IDProperty);
            set => SetValue(IDProperty, value);
        }
        #endregion

        #region Stroke
        public static readonly DependencyProperty StrokeProperty =
            DependencyProperty.Register(nameof(Stroke), typeof(string),
            typeof(ProbeSetupProgressBar), new PropertyMetadata("0"));

        public string Stroke
        {
            get => (string)GetValue(StrokeProperty);
            set => SetValue(StrokeProperty, value);
        }
        #endregion

        #region Status
        public static readonly DependencyProperty StatusProperty =
            DependencyProperty.Register(nameof(Status), typeof(string),
            typeof(ProbeSetupProgressBar), new PropertyMetadata(string.Empty));

        public string Status
        {
            get => (string)GetValue(StatusProperty);
            set => SetValue(StatusProperty, value);
        }
        #endregion

        // ---------------------------------------------------
        // Min/Max for probe
        // ---------------------------------------------------

        #region MinProbe
        public static readonly DependencyProperty MinProbeProperty =
            DependencyProperty.Register(nameof(MinProbe), typeof(double),
            typeof(ProbeSetupProgressBar), new PropertyMetadata(-1.173, OnProbeRangeChanged));

        public double MinProbe
        {
            get => (double)GetValue(MinProbeProperty);
            set => SetValue(MinProbeProperty, value);
        }
        #endregion

        #region MaxProbe
        public static readonly DependencyProperty MaxProbeProperty =
            DependencyProperty.Register(nameof(MaxProbe), typeof(double),
            typeof(ProbeSetupProgressBar), new PropertyMetadata(1.874, OnProbeRangeChanged));

        public double MaxProbe
        {
            get => (double)GetValue(MaxProbeProperty);
            set => SetValue(MaxProbeProperty, value);
        }
        #endregion

        private static void OnProbeRangeChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is ProbeSetupProgressBar control)
                control.UpdateMidLine();
        }

        // ---------------------------------------------------
        // Value property
        // ---------------------------------------------------

        #region Value
        public static readonly DependencyProperty ValueProperty =
            DependencyProperty.Register(nameof(Value), typeof(double),
            typeof(ProbeSetupProgressBar), new PropertyMetadata(0.0, OnValueChanged));

        public double Value
        {
            get => (double)GetValue(ValueProperty);
            set => SetValue(ValueProperty, value);
        }

        private static void OnValueChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is not ProbeSetupProgressBar control) return;

            double raw = (double)e.NewValue;
            control.UpdateBars(raw);
        }
        #endregion

        // ---------------------------------------------------
        // Update Bars based on Value
        // ---------------------------------------------------

        private void UpdateBars(double rawValue)
        {
            double min = MinProbe;
            double max = MaxProbe;
            double mid = 0.0;

            double fullRange = max - min;
            double shifted = rawValue - mid;

            // Half of total bar height = 125px
            double scaled = (shifted / fullRange) * 370;
            scaled = Math.Clamp(scaled, -140, 140);

            if (scaled < 0)
            {
                // Negative: fill downward
                AnimateRectangle(LeftBar, Math.Abs(scaled), growDown: true);
                AnimateRectangle(RightBar, 0, growDown: false);
            }
            else
            {
                // Positive: fill upward
                AnimateRectangle(RightBar, scaled, growDown: false);
                AnimateRectangle(LeftBar, 0, growDown: true);
            }
        }

        private void AnimateRectangle(Rectangle rect, double targetHeight, bool growDown)
        {
            double fromHeight = rect.Height;
            double fromTop = Canvas.GetTop(rect);
            double midY = 140;

            double toTop = growDown ? midY : midY - targetHeight;

            var heightAnim = new DoubleAnimation(fromHeight, targetHeight, TimeSpan.FromMilliseconds(120))
            {
                EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseOut }
            };

            var topAnim = new DoubleAnimation(fromTop, toTop, TimeSpan.FromMilliseconds(120))
            {
                EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseOut }
            };

            rect.BeginAnimation(Rectangle.HeightProperty, heightAnim);
            rect.BeginAnimation(Canvas.TopProperty, topAnim);
        }

        // ---------------------------------------------------
        // Mid line update (if Min/Max changes)
        // ---------------------------------------------------

        private void UpdateMidLine()
        {
            if (MaxProbe <= MinProbe) return;

            double mid = 0.0; // midpoint fixed at 0
            double normalized = (mid - MinProbe) / (MaxProbe - MinProbe);

            double targetY = 250 - (normalized * 250);
            double currentY = Canvas.GetTop(MidLine);

            var anim = new DoubleAnimation
            {
                From = double.IsNaN(currentY) ? targetY : currentY,
                To = targetY,
                Duration = TimeSpan.FromMilliseconds(200),
                EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseInOut }
            };

            MidLine.BeginAnimation(Canvas.TopProperty, anim);
        }
    }
}
