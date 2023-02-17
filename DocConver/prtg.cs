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

using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Listener;

namespace DocConver
{
    public partial class prtg : Form
    {
        private string connectionString;
        public prtg()
        {
            InitializeComponent();
        }

        private void RPTG_Load(object sender, EventArgs e)
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {
            connectionString = conn.Text;
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                try
                {
                    connection.Open();
                    MessageBox.Show("connection successful");
                    connection.Close();
                }
                catch (SqlException)
                {
                    MessageBox.Show("connection failed");
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
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
                    string pattern = @"Beschikbaarheid Stats: (\w+): (\d+ %)";
                    Match match = Regex.Match(input, pattern);
                    if (match.Success)
                    {
                        String qeury = "INSERT INTO prtg ([Beschikbaarheid Stats], Percentage) VALUES (@value1, @value2)";
                        using (SqlCommand command = new SqlCommand(qeury, connection))
                        {
                            command.Parameters.AddWithValue("@value1", match.Groups[1].Value);
                            command.Parameters.AddWithValue("@value2", match.Groups[2].Value);
                            command.ExecuteNonQuery();
                        }
                    }
                }
                connection.Close();
            }
            MessageBox.Show("data has been stored in dbo.ptrg");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SqlConnection conn = new SqlConnection(connectionString);
            conn.Open();

            String Query = "TRUNCATE TABLE prtg";
            SqlCommand cmd = new SqlCommand(Query, conn);
            cmd.ExecuteNonQuery();

            conn.Close();
            MessageBox.Show("data has been cleared");
        }

        private void button4_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "select file";
            openFileDialog.Filter = "pdf|*.pdf";
            openFileDialog.Multiselect = false;
            openFileDialog.ShowDialog();
            iPath.Text = openFileDialog.FileName;
        }

        private void button7_Click(object sender, EventArgs e)
        {
            string query = "CREATE TABLE [dbo].[prtg] (\r\n    [id]                    INT          IDENTITY (1, 1) NOT NULL,\r\n    [Beschikbaarheid Stats] VARCHAR (50) NULL,\r\n    [Percentage]            VARCHAR (50) NULL\r\n);\r\n";
            SqlConnection conn = new SqlConnection(connectionString);
            conn.Open();
            SqlCommand cmd = new SqlCommand(query, conn);
            cmd.ExecuteNonQuery();
            MessageBox.Show("Create successfully");
            conn.Close();
        }
    }
}
