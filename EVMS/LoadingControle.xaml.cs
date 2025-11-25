using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Animation;

namespace EVMS
{
    public partial class LoadingControle : UserControl
    {
        private Storyboard spinnerStoryboard;

        public LoadingControle()
        {
            InitializeComponent();
        }

        public void Start()
        {
            Visibility = Visibility.Visible;

            spinnerStoryboard = new Storyboard();

            DoubleAnimation rotationAnimation = new DoubleAnimation
            {
                From = 0,
                To = 360,
                Duration = TimeSpan.FromSeconds(1),
                RepeatBehavior = RepeatBehavior.Forever
            };

            Storyboard.SetTarget(rotationAnimation, SpinnerRotate);
            Storyboard.SetTargetProperty(rotationAnimation, new PropertyPath("Angle"));

            spinnerStoryboard.Children.Add(rotationAnimation);
            spinnerStoryboard.Begin();
        }

        public void Stop()
        {
            spinnerStoryboard?.Stop();
            Visibility = Visibility.Collapsed;
        }
    }
}
