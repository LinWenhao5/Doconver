using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading;

//execel dependency
//verantwoordelijk voor het lezen en schrijven van Excel-bestand
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;

using System.IO;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Configuration;
using iText.Layout.Font;
using iText.Kernel.Geom;

namespace DocConver
{
    public partial class servicedesk : Form
    {
        private string connectionString = ConfigurationManager.ConnectionStrings["localhost"].ConnectionString;

        public servicedesk()
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

        private void excelUploadentToDatabase(object sender, EventArgs e) //het uploaden van excel gegenes naar dbo.helpdesk
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                   Thread uploaden = new Thread(() => servicedesk.uploaden(connectionString, iPath.Text, 3, 4));
                   uploaden.Start();
                   MessageBox.Show("Het uploaden van de gegevens is beginnen...");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }


        private void emptyData(object sender, EventArgs e) //verwijder de data van dbo.helpDesk
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();

                String Query = "TRUNCATE TABLE helpDesk";
                SqlCommand cmd = new SqlCommand(Query, connection);
                cmd.ExecuteNonQuery();

                connection.Close();
                MessageBox.Show("data has been cleared");
            }
        }


        private void selectFile(object sender, EventArgs e) //Krijg de link van het excel bestand
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Title = "select file";
                openFileDialog.Filter = "excel|*.xlsx";
                openFileDialog.Multiselect = false;
                openFileDialog.ShowDialog();
                iPath.Text = openFileDialog.FileName;
            }   
        }

        public static void uploaden(string connectionString, string path, int x, int y)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = excelApp.Workbooks.Open(path);
            Excel.Worksheet worksheet = workbook.Worksheets[1];
            Excel.Range range = worksheet.UsedRange;

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();

                for (int row = 1; row <= range.Rows.Count; row++)
                {
                    for (int col = 1; col <= range.Columns.Count; col++)
                    {
                        //Sommige waarde in het excelblad zijn niet wat we willen
                        //dus die moet uitsluiten in de for-loop
                        if (row >= y && col == range.Columns.Count) // y.Value is het aantal rijen dat ik wil weglaten.
                        {
                            String query = "INSERT INTO helpDesk (#, Bedrijf, Title, [Naam verzoeker], Urgentie, [Tijd indienen verzoek], [Tijd sluiting], Status, Categorie, Subcategorie, [Type serviceverzoek] ,[Time to Respond], [Time to Repair], [Total Activities time])" +
                                "VALUES(@#, @bedrijf,@title,@naam,@urgentie,@tijd_ind_verzoek,@tijd_sluiting,@status,@cate,@subcate,@tpye_service,@time_to_resp,@time_to_repair,@total_act_time)";

                            using (SqlCommand command = new SqlCommand(query, connection))
                            {
                                //x.Value is het aantal kolom dat ik wil weglaten.
                                string[] value = new string[14];
                                DateTime? tijd_ind = null;
                                DateTime? tijd_slu = null;
                                for (int i = 0; i <= 13; i++)
                                {
                                    if (i == 4)
                                    {
                                        value[i] = range.Cells[row, x + i].value2;
                                        if (value[i] != null)
                                        {
                                            tijd_ind = DateTime.Parse(value[i]);
                                        }
                                    }
                                    else if (i == 5)
                                    {
                                        value[i] = range.Cells[row, x + i].value2;
                                        if (value[i] != null)
                                        {
                                            tijd_slu = DateTime.Parse(value[i]);
                                        }
                                    }
                                    else
                                    {
                                        value[i] = range.Cells[row, x + i].value2;
                                    }
                                }
                                command.Parameters.AddWithValue("@#", range.Cells[row, 1].value2);
                                command.Parameters.AddWithValue("@bedrijf", "" + value[0] + "");
                                command.Parameters.AddWithValue("@title", "" + value[1] + "");
                                command.Parameters.AddWithValue("@naam", "" + value[2] + "");
                                command.Parameters.AddWithValue("@urgentie", "" + value[3] + "");
                                command.Parameters.AddWithValue("@tijd_ind_verzoek", tijd_ind);
                                if (tijd_slu == null)
                                {
                                    command.Parameters.AddWithValue("@tijd_sluiting", DBNull.Value);
                                }
                                else
                                {
                                    command.Parameters.AddWithValue("@tijd_sluiting", tijd_slu);
                                }
                                command.Parameters.AddWithValue("@status", "" + value[6] + "");
                                command.Parameters.AddWithValue("@cate", "" + value[7] + "");
                                command.Parameters.AddWithValue("@subcate", "" + value[8] + "");
                                command.Parameters.AddWithValue("@tpye_service", "" + value[9] + "");
                                command.Parameters.AddWithValue("@time_to_resp", "" + value[10] + "");
                                command.Parameters.AddWithValue("@time_to_repair", "" + value[11] + "");
                                command.Parameters.AddWithValue("@total_act_time", "" + value[12] + "");
                                command.ExecuteNonQuery();
                            }
                        }

                    }
                }
                connection.Close();
                workbook.Close();
                excelApp.Quit();
                MessageBox.Show("data has been send");
            }
        }

    }
}
