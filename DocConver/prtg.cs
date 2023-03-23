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

//iText is een dependcy die inhoud van pdf kunt lezen
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using System.Security.Policy;
using System.Threading;

namespace DocConver
{
    public partial class prtg : Form
    {
        private string connectionString = ConfigurationManager.ConnectionStrings["localhost"].ConnectionString;

        public prtg()
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    InitializeComponent();
                    connection.Open();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void readData(object sender, EventArgs e) //get belangrijke data van pdf en opslaan in dbo prtg
        {
            Thread readPdf = new Thread(() => prtg.threadReadPdf(connectionString, iPath.Text));
            readPdf.Start();
        }

        private void emptyPrtg(object sender, EventArgs e)
        {

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();

                String Query = "TRUNCATE TABLE prtg";
                SqlCommand cmd = new SqlCommand(Query, connection);
                cmd.ExecuteNonQuery();

                connection.Close();
                MessageBox.Show("data has been cleared");
            }
        }

        private void selectFile(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "select file";
            openFileDialog.Filter = "pdf|*.pdf";
            openFileDialog.Multiselect = false;
            openFileDialog.ShowDialog();
            iPath.Text = openFileDialog.FileName;
        }

        public static void threadReadPdf(string connectionString, string file)
        {
            PdfReader pdfRead = new PdfReader(file);
            PdfDocument pdfDoc = new PdfDocument(pdfRead);
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                for (int page = 1; page <= pdfDoc.GetNumberOfPages(); page++) //pdf wordt per pagina gelezen
                {
                    ITextExtractionStrategy strategy = new SimpleTextExtractionStrategy();
                    //input bevat de inhoud van deze pagina
                    string input = PdfTextExtractor.GetTextFromPage(pdfDoc.GetPage(page), strategy);
                    Debug.WriteLine(input);
                    // regex worden hier gebruiken om belangrijke informatie uit de teskt halen
                    string pattern = @"De Periode van het rapport: (?<date>\d+\/\d+\/\d+)|Probe, Groep, Apparaat:(?<Component>.*)|Beschikbaarheid Stats: (?<status>\w+): (?<value>\d+\.\d+|\d+|)";
                    MatchCollection matches = Regex.Matches(input, pattern);
                    if (matches.Count == 3)
                    {
                        String qeury = "INSERT INTO prtg (Apparaat, [Beschikbaarheid Stats], Percentage, Datum) VALUES (@value1, @value2, @value3, @value4)";
                        using (SqlCommand command = new SqlCommand(qeury, connection))
                        {
                            //hier voeg de opgehaalde Group toe aan de query
                            command.Parameters.AddWithValue("@value1", matches[1].Groups["Component"].Value);
                            command.Parameters.AddWithValue("@value2", matches[2].Groups["status"].Value);
                            command.Parameters.AddWithValue("@value3", matches[2].Groups["value"].Value);
                            command.Parameters.AddWithValue("@value4", DateTime.Parse(matches[0].Groups["date"].Value));

                            command.ExecuteNonQuery();
                        }
                    }

                }
                connection.Close();
                MessageBox.Show("data has been save");
            }
        }
    }
}
