using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DocConver
{
    public partial class menu : Form
    {
        private Form servicedesk;
        private Form prtg;
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
    }
}
