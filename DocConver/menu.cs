using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Configuration;

namespace DocConver
{
    public partial class menu : Form
    {
        private string connectionString = ConfigurationManager.ConnectionStrings["localhost"].ConnectionString;
        private Form servicedesk;
        private Form prtg;
        private Form conn;
        public menu()
        {
            InitializeComponent();
        }

        // dit zijn vier knoppen van het menu
        private void openServicedesk(object sender, EventArgs e)
        {
            servicedesk = new servicedesk();
            servicedesk.Show();
        }

        private void openPrtg(object sender, EventArgs e)
        {
            prtg = new prtg();
            prtg.Show();
        }

        private void openConn(object sender, EventArgs e)
        {
            conn = new conn();
            conn.Show();
        }


        private void testConnection(object sender, EventArgs e)
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();
                    MessageBox.Show("connection successful");
                }
            } 
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

    }
}
