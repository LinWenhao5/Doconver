using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.Configuration;

using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Listener;

namespace DocConver
{
    public partial class prtg : Form
    {
        private string connectionString = ConfigurationManager.ConnectionStrings["localhost"].ConnectionString;
        public prtg()
        {
            InitializeComponent();
        }

        private void RPTG_Load(object sender, EventArgs e)
        {

        }

        private void read_data(object sender, EventArgs e)
        {
            string file = iPath.Text;
            PdfReader pdfRead = new PdfReader(file);
            PdfDocument pdfDoc = new PdfDocument(pdfRead);
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                for (int page = 1; page <= pdfDoc.GetNumberOfPages(); page++)
                {
                    ITextExtractionStrategy strategy = new SimpleTextExtractionStrategy();
                    string input = PdfTextExtractor.GetTextFromPage(pdfDoc.GetPage(page), strategy);
                    string pattern = @"Probe, Groep, Apparaat:(?<apparaat>.*)\nBeschikbaarheid Stats: (?<status>\w+): (?<percentage>\d*)";
                    Match match = Regex.Match(input, pattern);
                    if (match.Success)
                    {
                        String qeury = "INSERT INTO prtg (Apparaat, [Beschikbaarheid Stats], Percentage) VALUES (@value1, @value2, @value3)";
                        using (SqlCommand command = new SqlCommand(qeury, connection))
                        {
                            command.Parameters.AddWithValue("@value1", match.Groups["apparaat"].Value);
                            command.Parameters.AddWithValue("@value2", match.Groups["status"].Value);
                            command.Parameters.AddWithValue("@value3", match.Groups["percentage"].Value);
                            command.ExecuteNonQuery();
                        }
                    }
                }
                connection.Close();
            }
            MessageBox.Show("data has been stored in dbo.ptrg");
        }

        private void empty_prtg(object sender, EventArgs e)
        {
            SqlConnection conn = new SqlConnection(connectionString);
            conn.Open();

            String Query = "TRUNCATE TABLE prtg";
            SqlCommand cmd = new SqlCommand(Query, conn);
            cmd.ExecuteNonQuery();

            conn.Close();
            MessageBox.Show("data has been cleared");
        }

        private void select_file(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "select file";
            openFileDialog.Filter = "pdf|*.pdf";
            openFileDialog.Multiselect = false;
            openFileDialog.ShowDialog();
            iPath.Text = openFileDialog.FileName;
        }
    }
}
