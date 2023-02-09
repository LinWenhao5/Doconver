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

namespace DocConver
{
    public partial class Form1 : Form
    {

        public string connectionString = "Data Source=DESKTOP-6K3QQ5H\\SQLEXPRESS;Initial Catalog=test;Integrated Security=True;Connect Timeout=30;Encrypt=False;TrustServerCertificate=False;ApplicationIntent=ReadWrite;MultiSubnetFailover=False";
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string fileName = @"C:\Users\Lin\Documents\SLR_1675345541292.xlsx";
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = excelApp.Workbooks.Open(fileName);
            Excel.Worksheet worksheet = workbook.Worksheets[1];
            Excel.Range range = worksheet.UsedRange;


            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                for (int row = 1; row <= range.Rows.Count; row++)
                {
                    for (int col = 1; col <= range.Columns.Count; col++)
                    {
                        if (row > 4 && col == range.Columns.Count)
                        {
                            String query = "INSERT INTO helpDesk (Bedrijf, Title, [Naam verzoeker], Urgentie, [Tijd indienen verzoek], [Tijd sluiting], Status, Categorie, Subcategorie, [Type serviceverzoek] ,[Time to Respond], [Time to Repair], [Total Activities time])" +
                                "VALUES(@bedrijf,@title,@naam,@urgentie,@t_i_v,@tijd_s,@status,@c,@sc,@tpye_s,@t_o_res,@t_o_rep,@t_a_t)";

                            using (SqlCommand command = new SqlCommand(query, connection))
                            {
                                string col_one = range.Cells[row, 3].value2;
                                string col_two = range.Cells[row, 4].value2;
                                string col_three = range.Cells[row, 5].value2;
                                string col_four = range.Cells[row, 6].value2;
                                DateTime col_five = DateTime.Parse(range.Cells[row, 7].value2);
                                DateTime col_six = DateTime.Parse(range.Cells[row, 8].value2);
                                string col_seven = range.Cells[row, 9].value2;
                                string col_eight = range.Cells[row, 10].value2;
                                string col_nine = range.Cells[row, 11].value2;
                                string col_ten = range.Cells[row, 12].value2;
                                string col_eleven = range.Cells[row, 13].value2;
                                string col_twelve = range.Cells[row, 14].value2;
                                string col_thirteen = range.Cells[row, 15].value2;

                                command.Parameters.AddWithValue("@bedrijf", "" + col_one + "");
                                command.Parameters.AddWithValue("@title", "" + col_two + "");
                                command.Parameters.AddWithValue("@naam", "" + col_three + "");
                                command.Parameters.AddWithValue("@urgentie", "" + col_four + "");
                                command.Parameters.AddWithValue("@t_i_v",col_five);
                                command.Parameters.AddWithValue("@tijd_s",col_six);
                                command.Parameters.AddWithValue("@status", "" + col_seven + "");
                                command.Parameters.AddWithValue("@c", "" + col_eight + "");
                                command.Parameters.AddWithValue("@sc", "" + col_nine + "");
                                command.Parameters.AddWithValue("@tpye_s", "" + col_ten + "");
                                command.Parameters.AddWithValue("@t_o_res", "" + col_eleven + "");
                                command.Parameters.AddWithValue("@t_o_rep", "" + col_twelve + "");
                                command.Parameters.AddWithValue("@t_a_t", "" + col_thirteen + "");
                                //MessageBox.Show(col_five.ToString());
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

        private void button2_Click(object sender, EventArgs e)
        {
            SqlConnection conn = new SqlConnection(connectionString);
            conn.Open();

            String Query = "TRUNCATE TABLE helpDesk";
            SqlCommand cmd = new SqlCommand(Query, conn);
            cmd.ExecuteNonQuery();

            conn.Close();
            MessageBox.Show("data has been cleared");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string filePath = "C:\\Users\\Lin\\Downloads\\result.xlsx";
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = excelApp.Workbooks.Add();
            Excel.Worksheet worksheet = workbook.Worksheets[1];
            Excel.Range range = worksheet.UsedRange;

            string bd =  beginDatum.Value.ToString("yyyy-MM-dd");
            string ed = eindDatum.Value.ToString("yyyy-MM-dd");

            range.Cells[1, 1].value = "Tijdsperiode";
            range.Cells[1, 2].value = "Gesloten Changes";
            range.Cells[1, 3].value = "Gesloten Incidents";
            range.Cells[1, 4].value = "Gesloten Problems";
            range.Cells[1, 5].value = "Gesloten Service Requests";
            range.Cells[2, 1].value = bd + " tot " + ed;

            SqlConnection conn = new SqlConnection(connectionString);
            conn.Open();
            string[] qeurys = {
                "SELECT COUNT(*) FROM helpDesk WHERE [Type serviceverzoek]='Change' AND Status='Gesloten' AND [Tijd sluiting] BETWEEN '"+bd+"' AND '"+ed+"'",
                "SELECT COUNT(*) FROM helpDesk WHERE [Type serviceverzoek]='Incident' AND Status='Gesloten' AND [Tijd sluiting] BETWEEN '"+bd+"' AND '"+ed+"'",
                "SELECT COUNT(*) FROM helpDesk WHERE [Type serviceverzoek]='Problem' AND Status='Gesloten' AND [Tijd sluiting] BETWEEN '"+bd+"' AND '"+ed+"'",
                "SELECT COUNT(*) FROM helpDesk WHERE [Type serviceverzoek]='Request' AND Status='Gesloten' AND [Tijd sluiting] BETWEEN '"+bd+"' AND '"+ed+"'"
            };
            int[] result = new int[4];
            for (int i = 0; i < qeurys.Length; i++)
            {
                Debug.WriteLine(qeurys[i]);
                SqlCommand cmd = new SqlCommand(qeurys[i], conn);
                Int32 count = (Int32)(cmd.ExecuteScalar());
                range.Cells[2, i+2].value = count;
            }

            workbook.SaveAs("C:\\Users\\Lin\\Downloads\\result.xlsx");
            conn.Close();
            workbook.Close();
            excelApp.Quit();

            MessageBox.Show("Downloads path: " + filePath);

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

    }
}
