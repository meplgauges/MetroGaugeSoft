using Microsoft.Data.SqlClient;
using System.Configuration;
using System.Windows;

namespace EVMS
{
    public partial class Login_Page : Window
    {

        private readonly string connectionString;

        public Login_Page()
        {
            InitializeComponent();
            connectionString = ConfigurationManager.ConnectionStrings["EVMSDb"].ConnectionString;

        }

        private void BtnLogin_Click(object sender, RoutedEventArgs e)
        {
            string username = txtUsername.Text.Trim();
            string password = txtPassword.Password;

            string query = "SELECT UserType FROM Users WHERE UserID = @UserID AND Password = @Password";

            using (SqlConnection con = new SqlConnection(connectionString))
            {
                con.Open();

                using (SqlCommand cmd = new SqlCommand(query, con))
                {
                    cmd.Parameters.AddWithValue("@UserID", username);
                    cmd.Parameters.AddWithValue("@Password", password);

                    object result = cmd.ExecuteScalar();

                    if (result != null)
                    {
                        SessionManager.IsAuthenticated = true;
                        SessionManager.UserType = result.ToString();
                        SessionManager.UserID = username;

                        this.DialogResult = true;
                        this.Close();
                    }
                    else
                    {
                        SessionManager.IsAuthenticated = false;
                        MessageBox.Show("Invalid Username or Password!");
                    }
                }
            }
        }

        public static class SessionManager
        {
            public static string UserID { get; set; }
            public static string UserType { get; set; }
            public static bool IsAuthenticated { get; set; }
        }


        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            SessionManager.IsAuthenticated = false;
            this.DialogResult = false;
            this.Close();
        }
    }
}
