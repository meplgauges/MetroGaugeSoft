using EVMS.Service;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.Windows.Threading;

namespace EVMS
{
    /// <summary>
    /// Interaction logic for SplashScreenPage.xaml
    /// </summary>
    public partial class SplashScreenPage : Window
    {
        private readonly string[] _loadingSteps;
        private int _currentStepIndex = 0;
        private readonly DispatcherTimer _stepTimer;

        public SplashScreenPage()
        {
            InitializeComponent();

            _loadingSteps = new string[]
            {
                "Initializing Engine...",
                "Loading Assets...",
                "Connecting Services...",
                "Finalizing..."
            };

            _stepTimer = new DispatcherTimer
            {
                Interval = TimeSpan.FromMilliseconds(800)
            };
            _stepTimer.Tick += OnStepTimerTick;
            _stepTimer.Start();

            ProgressBar.Value = 0;
            LoadingText.Text = "Starting...";
        }

        private async void OnStepTimerTick(object? sender, EventArgs? e)
        {
            if (_currentStepIndex < _loadingSteps.Length)
            {
                UpdateProgress(_loadingSteps[_currentStepIndex]);
                _currentStepIndex++;
            }
            else
            {
                _stepTimer.Stop();
                await CompleteInitializationAsync();
            }
        }

        private void UpdateProgress(string message)
        {
            LoadingText.Text = message;
            ProgressBar.Value = ((_currentStepIndex + 2) * 100) / _loadingSteps.Length;
        }

        private async Task CompleteInitializationAsync()
        {
            LoadingText.Text = "Launching...";
            ProgressBar.Value = 100;

            await Task.Delay(1000); // Optional visual delay

            await Task.Run(() =>
            {
                // Simulate startup tasks (e.g., license check, config loading)
                Task.Delay(500).Wait();
            });

            LaunchMainWindow();
        }

       


        private void LaunchMainWindow()
        {
            var mainWindow = new MainWindow();
            mainWindow.Show();
            this.Close();
        }
    }
}