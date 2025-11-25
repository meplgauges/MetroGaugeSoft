using System;
using System.Threading;
using System.Windows;

namespace EVMS
{
    public partial class App : Application
    {
        private static Mutex _mutex;

        private void Application_Startup(object sender, StartupEventArgs e)
        {
            bool isNewInstance;
            string mutexName = "EVMSUniqueApplicationMutex";

            _mutex = new Mutex(true, mutexName, out isNewInstance);

            if (!isNewInstance)
            {
                MessageBox.Show("The application is already running.", "Instance Running", MessageBoxButton.OK, MessageBoxImage.Information);
                Current.Shutdown();
                return;
            }

            var splash = new SplashScreenPage();
            splash.Show();
        }

        protected override void OnExit(ExitEventArgs e)
        {
            if (_mutex != null)
            {
                _mutex.ReleaseMutex();
                _mutex.Dispose();
            }
            base.OnExit(e);
        }


    }
}
