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
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;
using System.IO;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Configuration;

namespace DocConver
{
    public partial class servicedesk : Form
    {
        private string connectionString = ConfigurationManager.ConnectionStrings["localhost"].ConnectionString;
        public servicedesk()
        {
            InitializeComponent();
        }

        private void excelUploadentToDatabase(object sender, EventArgs e)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = excelApp.Workbooks.Open(iPath.Text);
            Excel.Worksheet worksheet = workbook.Worksheets[1];
            Excel.Range range = worksheet.UsedRange;


            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                for (int row = 1; row <= range.Rows.Count; row++)
                {
                    for (int col = 1; col <= range.Columns.Count; col++)
                    {
                        if (row >= y.Value && col == range.Columns.Count)
                        {
                            String query = "INSERT INTO helpDesk (Bedrijf, Title, [Naam verzoeker], Urgentie, [Tijd indienen verzoek], [Tijd sluiting], Status, Categorie, Subcategorie, [Type serviceverzoek] ,[Time to Respond], [Time to Repair], [Total Activities time])" +
                                "VALUES(@bedrijf,@title,@naam,@urgentie,@tijd_ind_verzoek,@tijd_sluiting,@status,@cate,@subcate,@tpye_service,@time_to_resp,@time_to_repair,@total_act_time)";

                            using (SqlCommand command = new SqlCommand(query, connection))
                            {
                                string col_one = range.Cells[row, x.Value].value2;
                                string col_two = range.Cells[row, x.Value + 1].value2;
                                string col_three = range.Cells[row, x.Value + 2].value2;
                                string col_four = range.Cells[row, x.Value + 3].value2;
                                DateTime col_five = DateTime.Parse(range.Cells[row, x.Value + 4].value2);
                                DateTime col_six = DateTime.MaxValue;
                                if (range.Cells[row, x.Value + 5].value2 != null)
                                {
                                    col_six = DateTime.Parse(range.Cells[row, x.Value + 5].value2);
                                }
                                string col_seven = range.Cells[row, x.Value + 6].value2;
                                string col_eight = range.Cells[row, x.Value + 7].value2;
                                string col_nine = range.Cells[row, x.Value + 8].value2;
                                string col_ten = range.Cells[row, x.Value + 9].value2;
                                string col_eleven = range.Cells[row, x.Value + 10].value2;
                                string col_twelve = range.Cells[row, x.Value + 11].value2;
                                string col_thirteen = range.Cells[row, x.Value + 12].value2;

                                command.Parameters.AddWithValue("@bedrijf", "" + col_one + "");
                                command.Parameters.AddWithValue("@title", "" + col_two + "");
                                command.Parameters.AddWithValue("@naam", "" + col_three + "");
                                command.Parameters.AddWithValue("@urgentie", "" + col_four + "");
                                command.Parameters.AddWithValue("@tijd_ind_verzoek", col_five);
                                command.Parameters.AddWithValue("@tijd_sluiting", col_six);
                                command.Parameters.AddWithValue("@status", "" + col_seven + "");
                                command.Parameters.AddWithValue("@cate", "" + col_eight + "");
                                command.Parameters.AddWithValue("@subcate", "" + col_nine + "");
                                command.Parameters.AddWithValue("@tpye_service", "" + col_ten + "");
                                command.Parameters.AddWithValue("@time_to_resp", "" + col_eleven + "");
                                command.Parameters.AddWithValue("@time_to_repair", "" + col_twelve + "");
                                command.Parameters.AddWithValue("@total_act_time", "" + col_thirteen + "");
                                command.ExecuteNonQuery();
                            }
                        }

                    }
                }
                connection.Close();
            }
            workbook.Close();
            excelApp.Quit();
            MessageBox.Show("data has been send");
        }

        private void emptyData(object sender, EventArgs e)
        {
            SqlConnection conn = new SqlConnection(connectionString);
            conn.Open();

            String Query = "TRUNCATE TABLE helpDesk";
            SqlCommand cmd = new SqlCommand(Query, conn);
            cmd.ExecuteNonQuery();

            conn.Close();
            MessageBox.Show("data has been cleared");
        }


        private void selectFile(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "select file";
            openFileDialog.Filter = "excel|*.xlsx";
            openFileDialog.Multiselect = false;
            openFileDialog.ShowDialog();
            iPath.Text = openFileDialog.FileName;
        }
    }
}
