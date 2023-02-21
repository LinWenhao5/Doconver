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

        private void button1_Click(object sender, EventArgs e)
        {
            servicedesk = new servicedesk();
            servicedesk.Show();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            prtg = new prtg();
            prtg.Show();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            conn = new conn();
            conn.Show();
        }


        private void button4_Click(object sender, EventArgs e)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                try
                {
                    MessageBox.Show("connection successful");
                    connection.Close();
                }
                catch (SqlException ex)
                {
                    MessageBox.Show(ex.Message.ToString());
                }
            }
        }
    }
}
