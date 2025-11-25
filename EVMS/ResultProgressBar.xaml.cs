using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Shapes;

namespace EVMS
{
    public partial class ResultProgressBar : UserControl
    {
        // Dependency Properties
        public static readonly DependencyProperty MinValueProperty =
            DependencyProperty.Register(nameof(MinValue), typeof(double), typeof(ResultProgressBar),
                new PropertyMetadata(0.0, OnMinMaxChanged));

        public double MinValue
        {
            get => (double)GetValue(MinValueProperty);
            set => SetValue(MinValueProperty, value);
        }

        public static readonly DependencyProperty MaxValueProperty =
            DependencyProperty.Register(nameof(MaxValue), typeof(double), typeof(ResultProgressBar),
                new PropertyMetadata(1.0, OnMinMaxChanged));

        public double MaxValue
        {
            get => (double)GetValue(MaxValueProperty);
            set => SetValue(MaxValueProperty, value);
        }

        public static readonly DependencyProperty MeanValueProperty =
            DependencyProperty.Register(nameof(MeanValue), typeof(double), typeof(ResultProgressBar),
                new PropertyMetadata(0.5, OnMeanValueChanged));

        public double MeanValue
        {
            get => (double)GetValue(MeanValueProperty);
            set => SetValue(MeanValueProperty, value);
        }

        public static readonly DependencyProperty ParameterNameProperty =
            DependencyProperty.Register(nameof(ParameterName), typeof(string), typeof(ResultProgressBar),
                new PropertyMetadata(string.Empty));

        public string ParameterName
        {
            get => (string)GetValue(ParameterNameProperty);
            set => SetValue(ParameterNameProperty, value);
        }

        public static readonly DependencyProperty ValueProperty =
            DependencyProperty.Register(nameof(Value), typeof(double), typeof(ResultProgressBar),
                new PropertyMetadata(0.0, OnValueChanged));

        public double Value
        {
            get => (double)GetValue(ValueProperty);
            set => SetValue(ValueProperty, value);
        }

        //public static readonly DependencyProperty IsOkProperty =
        //    DependencyProperty.Register(nameof(IsOk), typeof(bool?), typeof(ResultProgressBar),
        //        new PropertyMetadata(null, OnIsOkChanged));

        //public bool? IsOk
        //{
        //    get => (bool?)GetValue(IsOkProperty);
        //    set => SetValue(IsOkProperty, value);
        //}

        // Geometry
        private double cx, cy, radius;

        public ResultProgressBar()
        {
            InitializeComponent();
            Loaded += ResultProgressBar_Loaded;
            SizeChanged += ResultProgressBar_SizeChanged;
        }

        private void ResultProgressBar_Loaded(object sender, RoutedEventArgs e)
        {
            ComputeGeometry();
            Redraw();
            PositionStaticElements();

            // Initialize display
            CenterValueText.Text = "0.000";
            UpdateArcColor();
            UpdateNeedle(MinValue);
        }

        private void ResultProgressBar_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            ComputeGeometry();
            Redraw();
        }

        private static void OnMinMaxChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is ResultProgressBar ctrl)
                ctrl.Redraw();
        }

        private static void OnMeanValueChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is ResultProgressBar ctrl)
                ctrl.UpdateValue((double)e.NewValue);
        }

        private static void OnValueChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is ResultProgressBar ctrl)
            {
                double newValue = (double)e.NewValue;
                ctrl.UpdateValue(newValue);
            }
        }

        //private static void OnIsOkChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        //{
        //    if (d is ResultProgressBar ctrl)
        //    {
        //        ctrl.UpdateArcColor();
        //    }
        //}

        private void ComputeGeometry()
        {
            double w = GaugeCanvas.ActualWidth;
            double h = GaugeCanvas.ActualHeight;
            if (w <= 0 || h <= 0) return;
            cx = w / 2.0;
            cy = h * 0.70;
            radius = Math.Min(w, h) * 0.45;
        }

        private void Redraw()
        {
            ArcsLayer.Children.Clear();
            DrawArc(MinValue, MaxValue, Brushes.LightGray);
            DrawTicksAndLabels();
            PositionStaticElements();
        }

        private void UpdateArcColor()
        {
            ArcsLayer.Children.Clear();

            // Always draw background arc in light gray
            DrawArc(MinValue, MaxValue, Brushes.LightGray);

            // If Value is (close to) zero, skip drawing colored arc
            if (Math.Abs(Value) < 0.0001)
            {
                return; // No colored overlay
            }

            Brush brush;
            if (Value < MinValue || Value > MaxValue)
                brush = Brushes.IndianRed;
            else
                brush = Brushes.LimeGreen;

            DrawArc(MinValue, MaxValue, brush);
        }





        private void DrawArc(double fromVal, double toVal, Brush brush)
        {
            double startTheta = ValueToTheta(fromVal);
            double endTheta = ValueToTheta(toVal);

            Point startPt = PolarToCartesian(cx, cy, radius, startTheta);
            Point endPt = PolarToCartesian(cx, cy, radius, endTheta);

            var fig = new PathFigure { StartPoint = startPt };
            var arcSeg = new ArcSegment
            {
                Point = endPt,
                Size = new Size(radius, radius),
                SweepDirection = SweepDirection.Clockwise,
                IsLargeArc = true
            };
            fig.Segments.Add(arcSeg);

            var geo = new PathGeometry();
            geo.Figures.Add(fig);

            var path = new Path
            {
                Data = geo,
                Stroke = brush,
                StrokeThickness = radius * 0.12,
                StrokeStartLineCap = PenLineCap.Round,
                StrokeEndLineCap = PenLineCap.Round
            };

            ArcsLayer.Children.Add(path);
        }

        private void DrawTicksAndLabels()
        {
            TicksLayer.Children.Clear();
            if (MaxValue <= MinValue) return;

            int majorDivisions = 6;
            int minorDivisions = 4;
            int totalTicks = (majorDivisions - 1) * minorDivisions + majorDivisions;
            double step = (MaxValue - MinValue) / (totalTicks - 1);
            double arcThickness = radius * 0.18;
            double tickStartR = radius - (arcThickness / 2.0) - 8;
            double tickEndR_major = tickStartR - radius * 0.08;
            double tickEndR_minor = tickStartR - radius * 0.04;

            int midMajorTickIndex = (totalTicks - 1) / 2;

            for (int i = 0; i < totalTicks; i++)
            {
                double val = (i == 0) ? MinValue : (i == totalTicks - 1) ? MaxValue : MinValue + i * step;
                double theta = ValueToTheta(val);
                double rad = theta * Math.PI / 180.0;
                bool isMajor = (i % minorDivisions == 0) || i == 0 || i == totalTicks - 1;

                Brush tickBrush = Brushes.Black;
                if (i == 0) tickBrush = Brushes.Blue;
                else if (i == midMajorTickIndex) tickBrush = Brushes.Green;
                else if (i == totalTicks - 1) tickBrush = Brushes.Red;

                var tick = new Line
                {
                    X1 = cx + tickStartR * Math.Cos(rad),
                    Y1 = cy - tickStartR * Math.Sin(rad),
                    X2 = cx + (isMajor ? tickEndR_major : tickEndR_minor) * Math.Cos(rad),
                    Y2 = cy - (isMajor ? tickEndR_major : tickEndR_minor) * Math.Sin(rad),
                    Stroke = tickBrush,
                    StrokeThickness = isMajor ? 2 : 1
                };
                TicksLayer.Children.Add(tick);

                if (isMajor)
                {
                    double fontSize = radius * 0.07;
                    fontSize = Math.Max(fontSize, 15); // ensure font size is at least 10
                    var tb = new TextBlock
                    {
                        FontSize = fontSize,
                        FontWeight = FontWeights.Bold,
                        Foreground = tickBrush,
                        Text = val.ToString("0.000")
                    };
                    double labelRadius = tickEndR_major - radius * 0.06;
                    double lx = cx + labelRadius * Math.Cos(rad);
                    double ly = cy - labelRadius * Math.Sin(rad);
                    tb.Measure(new Size(double.PositiveInfinity, double.PositiveInfinity));
                    Canvas.SetLeft(tb, lx - tb.DesiredSize.Width / 2);
                    Canvas.SetTop(tb, ly - tb.DesiredSize.Height / 2);
                    TicksLayer.Children.Add(tb);
                }
            }
        }

        private void PositionStaticElements()
        {
            if (radius <= 0) return;

            double needleLength = radius * 0.75;
            Needle.X1 = cx;
            Needle.Y1 = cy;
            Needle.X2 = cx;
            Needle.Y2 = cy - needleLength;
            NeedleRotate.CenterX = cx;
            NeedleRotate.CenterY = cy;

            HubContainer.Width = radius * 2.0;
            HubContainer.Height = radius * 0.3;
            Canvas.SetLeft(HubContainer, cx - HubContainer.Width / 2);
            Canvas.SetTop(HubContainer, cy + radius * 0.1);  // slightly below the center line (cy)

            double hubSize = radius * 0.35;
            Canvas.SetLeft(CenterValueText, cx - hubSize * 0.45);
            Canvas.SetTop(CenterValueText, cy - hubSize * 0.55);

            LabelID.FontSize = 20;
            LabelID.FontWeight = FontWeights.Bold;
            LabelID.Width = HubContainer.Width;
            LabelID.TextAlignment = TextAlignment.Center;

            // Position LabelID at very top of canvas, centered horizontally
            Canvas.SetLeft(LabelID, cx - LabelID.Width / 2);
            Canvas.SetTop(LabelID, 25); // 25 px margin from top edge

            // Add small round needle hub design at needle base
            // Remove any existing needle hub ellipse to avoid duplicates
            for (int i = GaugeCanvas.Children.Count - 1; i >= 0; i--)
            {
                if (GaugeCanvas.Children[i] is Ellipse e && e.Tag?.ToString() == "NeedleHub")
                {
                    GaugeCanvas.Children.RemoveAt(i);
                }
            }

            var needleHub = new Ellipse
            {
                Width = radius * 0.1,
                Height = radius * 0.1,
                Fill = Brushes.Black,
                Stroke = Brushes.Gray,
                StrokeThickness = 1,
                IsHitTestVisible = false,
                Tag = "NeedleHub"
            };

            Canvas.SetLeft(needleHub, cx - needleHub.Width / 2);
            Canvas.SetTop(needleHub, cy - needleHub.Height / 2);

            GaugeCanvas.Children.Add(needleHub);
        }





        private void UpdateNeedle(double value)
        {
            if (MaxValue == MinValue) // Prevent division by zero
            {
                return; // or handle default angle, e.g., 0 degrees
            }

            double center = (MinValue + MaxValue) / 2.0;
            double halfSpan = (MaxValue - MinValue) / 2.0;
            double normalized = (value - center) / halfSpan;
            double targetAngle = Math.Max(-90, Math.Min(90, normalized * 90));

            // Check for NaN before animating
            if (double.IsNaN(targetAngle))
            {
                targetAngle = 0; // fallback angle
            }

            var anim = new DoubleAnimation
            {
                To = targetAngle,
                Duration = TimeSpan.FromMilliseconds(400),
                EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseOut }
            };

            NeedleRotate.BeginAnimation(RotateTransform.AngleProperty, anim);
        }


        // **Main method to update value + OK status**
        public void UpdateValue(double value, bool? isOk = null)
        {
            Value = value;

            // Set HubContainer background based on OK status
            if (StatusRectangle != null)
            {
                if (isOk.HasValue)
                {
                    StatusRectangle.Fill = isOk.Value ? Brushes.LimeGreen : Brushes.IndianRed;
                }
                else
                {
                    StatusRectangle.Fill = Brushes.Black; // default black fill
                }
            }

            if (CenterValueText != null)
                CenterValueText.Text = value.ToString("0.000");

            UpdateArcColor();
            UpdateNeedle(value);
        }



        private double ValueToTheta(double value)
        {
            return 180.0 - ((value - MinValue) / (MaxValue - MinValue)) * 180.0;
        }

        private static Point PolarToCartesian(double cx, double cy, double r, double angleDeg)
        {
            double rad = angleDeg * Math.PI / 180.0;
            return new Point(cx + r * Math.Cos(rad), cy - r * Math.Sin(rad));
        }
    }
}
