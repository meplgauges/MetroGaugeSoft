using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using static EVMS.Login_Page;

namespace EVMS
{
    public partial class Dashboard : UserControl
    {
        public Grid MainContentGrid { get; set; }
        public event Action<string> StatusMessageChanged;
        public event Action<string> DashboardStatusMessageChanged;



        public Dashboard()
        {
            InitializeComponent();
            this.Loaded += Dashboard_Loaded;
        }

        private void Dashboard_Loaded(object sender, RoutedEventArgs e)
        {
            SetButtonAccessByRole();
        }

        private void SetButtonAccessByRole()
        {
            // ✅ Protect against null or empty user type
            string userType = SessionManager.UserType;

            if (string.Equals(userType, "Admin", StringComparison.OrdinalIgnoreCase))
            {
                // Admin → enable all buttons
                SetAllButtonsEnabled(true);
            }
            else
            {
                // Non-admin users → disable only restricted buttons
                SetButtonsEnabled(new List<string>
        {
            "Login/Setup",
            "Part Manager",
            "Settings",
            "Calculation"
            // Add more if needed
        }, false);
            }
        }


        private void UpdateStatus(string message)
        {
            StatusMessageChanged?.Invoke(message);
        }

        private void SetAllButtonsEnabled(bool enabled)
        {
            foreach (var btn in FindButtons(this))
            {
                btn.IsEnabled = enabled;
            }
        }

        private void SetButtonsEnabled(List<string> buttonContents, bool enabled)
        {
            foreach (var btn in FindButtons(this))
            {
                if (btn.Content is string content && buttonContents.Contains(content))
                {
                    btn.IsEnabled = enabled;
                }
            }
        }

        private IEnumerable<Button> FindButtons(DependencyObject parent)
        {
            if (parent == null)
                yield break;

            int count = VisualTreeHelper.GetChildrenCount(parent);

            for (int i = 0; i < count; i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);

                if (child is Button btn)
                    yield return btn;

                foreach (var childOfChild in FindButtons(child))
                    yield return childOfChild;
            }
        }

        private void DashboardCard_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button button && button.Content is string key)
            {
                UserControl pageToShow = null;

                switch (key)
                {
                    case "Measurement":
                        // ✅ Create EntryPage
                        var entryPage = new EntryPage();
                        entryPage.StartClicked += EntryPage_StartClicked;

                        // ✅ Assign MainContentGrid (reference from MainWindow)
                        var currentWindowForEntry = Window.GetWindow(this);
                        entryPage.MainContentGrid = currentWindowForEntry?.FindName("MainContentGrid") as Grid;

                        pageToShow = entryPage;
                        break;

                    case "Probe Setup":
                        pageToShow = new ProbeSetupPage();
                        break;

                    case "Master Reading":
                        pageToShow = new MasterReadingPage();
                        break;

                    case "Login/Setup":
                        pageToShow = new AdminControlePage();
                        break;

                    case "Report Graph":
                        pageToShow = new Report_GraphPage();
                        break;

                    case "Report View":
                        pageToShow = new Report_View_Page();
                        break;

                    case "Repeatability":
                        pageToShow = new Repeatbilty_Page();
                        break;

                    case "IO Control":
                        pageToShow = new IO_Controle_page();
                        break;

                    case "Settings":
                        pageToShow = new SettingsPage();
                        break;

                    case "Probe Install":
                        pageToShow = new ProbeInstallPage();
                        break;

                    case "Part Config":
                        pageToShow = new PartConfig();
                        break;

                    case "Part Manager":
                        pageToShow = new Part_Manager();
                        break;

                    case "Calculation":
                        pageToShow = new Calculation_Modification();
                        break;
                }

                if (pageToShow != null)
                {
                    var currentWindow = Window.GetWindow(this);
                    var mainContentGrid = currentWindow?.FindName("MainContentGrid") as Grid;

                    if (mainContentGrid != null)
                    {
                        mainContentGrid.Children.Clear();
                        pageToShow.HorizontalAlignment = HorizontalAlignment.Stretch;
                        pageToShow.VerticalAlignment = VerticalAlignment.Stretch;
                        mainContentGrid.Children.Add(pageToShow);
                    }
                    else
                    {
                        MessageBox.Show("MainContentGrid not found in MainWindow.");
                    }
                }
            }
        }

        // ✅ When EntryPage raises StartClicked → open ResultPage
        private void EntryPage_StartClicked(object sender, StartClickedEventArgs e)
        {
            Grid mainContentGrid = null;

            if (sender is EntryPage entryPage && entryPage.MainContentGrid != null)
            {
                mainContentGrid = entryPage.MainContentGrid;
            }
            else
            {
                var currentWindow = Window.GetWindow(this);
                mainContentGrid = currentWindow?.FindName("MainContentGrid") as Grid;
            }

            if (mainContentGrid != null)
            {
                mainContentGrid.Children.Clear();

                var resultPage = new ResultPage(e.Model, e.LotNo, e.UserId)
                {
                    HorizontalAlignment = HorizontalAlignment.Stretch,
                    VerticalAlignment = VerticalAlignment.Stretch
                };

                mainContentGrid.Children.Add(resultPage);

                // ✅ Hook event AFTER adding to grid
                resultPage.StatusMessageChanged += (message) =>
                {
                    // Forward the message up to MainWindow through Dashboard’s own event
                    DashboardStatusMessageChanged?.Invoke(message);
                };


            }
            else
            {
                MessageBox.Show("MainContentGrid not found in MainWindow. Cannot open ResultPage.");
            }
        }




    }

}
