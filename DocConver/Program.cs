using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DocConver
{
    internal static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// C:\Users\Lin\Documents\SLR_1675345541292.xlsx
        /// C:\Users\Lin\Downloads\result.xlsx
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }
}
