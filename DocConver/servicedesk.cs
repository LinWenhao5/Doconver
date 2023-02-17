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

namespace DocConver
{
    public partial class servicedesk : Form
    {
        private string connectionString;
        public servicedesk()
        {
            InitializeComponent();
        }

        private void connect_string_Click(object sender, EventArgs e)
        {
            connectionString = conn.Text;
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                try
                {
                    connection.Open();
                    MessageBox.Show("connection successful");
                    connection.Close();
                } catch (SqlException)
                {
                    MessageBox.Show("connection failed");
                }
            }
        }

        private void excel_uploaden_to_database_Click(object sender, EventArgs e)
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
                                string col_two = range.Cells[row, x.Value+1].value2;
                                string col_three = range.Cells[row, x.Value+2].value2;
                                string col_four = range.Cells[row, x.Value+3].value2;
                                DateTime col_five = DateTime.Parse(range.Cells[row, x.Value + 4].value2);
                                DateTime col_six = DateTime.MaxValue;
                                if (range.Cells[row, x.Value + 5].value2 != null)
                                {
                                    col_six = DateTime.Parse(range.Cells[row, x.Value + 5].value2);
                                }
                                string col_seven = range.Cells[row, x.Value+6].value2;
                                string col_eight = range.Cells[row, x.Value+7].value2;
                                string col_nine = range.Cells[row, x.Value+8].value2;
                                string col_ten = range.Cells[row, x.Value+9].value2;
                                string col_eleven = range.Cells[row, x.Value+10].value2;
                                string col_twelve = range.Cells[row, x.Value+11].value2;
                                string col_thirteen = range.Cells[row, x.Value+12].value2;

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

        private void empty_database_Click(object sender, EventArgs e)
        {
            SqlConnection conn = new SqlConnection(connectionString);
            conn.Open();

            String Query = "TRUNCATE TABLE data";
            SqlCommand cmd = new SqlCommand(Query, conn);
            cmd.ExecuteNonQuery();

            conn.Close();
            MessageBox.Show("data has been cleared");
        }

        private void export_data_to_excel_Click(object sender, EventArgs e)
        {
            string bd = beginDatum.Value.ToString("yyyy-MM-dd");
            string ed = eindDatum.Value.ToString("yyyy-MM-dd");

            SqlConnection conn = new SqlConnection(connectionString);
            conn.Open();
            string[] qeurys = {
                "SELECT COUNT(*) FROM helpDesk WHERE [Type serviceverzoek]='Change' AND Status='Gesloten' AND [Tijd sluiting] BETWEEN '"+bd+"' AND '"+ed+"'",
                "SELECT COUNT(*) FROM helpDesk WHERE [Type serviceverzoek]='Incident' AND Status='Gesloten' AND [Tijd sluiting] BETWEEN '"+bd+"' AND '"+ed+"'",
                "SELECT COUNT(*) FROM helpDesk WHERE [Type serviceverzoek]='Problem' AND Status='Gesloten' AND [Tijd sluiting] BETWEEN '"+bd+"' AND '"+ed+"'",
                "SELECT COUNT(*) FROM helpDesk WHERE [Type serviceverzoek]='Request' AND Status='Gesloten' AND [Tijd sluiting] BETWEEN '"+bd+"' AND '"+ed+"'",
                "SELECT COUNT(*) FROM helpDesk WHERE [Type serviceverzoek]='Change' AND Status='Pending' AND [Tijd indienen verzoek] BETWEEN '"+bd+"' AND '"+ed+"'",
                "SELECT COUNT(*) FROM helpDesk WHERE [Type serviceverzoek]='Incident' AND Status='Pending' AND [Tijd indienen verzoek] BETWEEN '"+bd+"' AND '"+ed+"'",
                "SELECT COUNT(*) FROM helpDesk WHERE [Type serviceverzoek]='Problem' AND Status='Pending' AND [Tijd indienen verzoek] BETWEEN '"+bd+"' AND '"+ed+"'",
                "SELECT COUNT(*) FROM helpDesk WHERE [Type serviceverzoek]='Request' AND Status='Pending' AND [Tijd indienen verzoek] BETWEEN '"+bd+"' AND '"+ed+"'"
            };

            int[] data = new int[8];
            for (int i = 0; i < qeurys.Length; i++)
            {
                SqlCommand cmd = new SqlCommand(qeurys[i], conn);
                Int32 count = (Int32)(cmd.ExecuteScalar());
                Debug.WriteLine(qeurys[i]);
                data[i] = count;
            }
            string query = "INSERT INTO data (Tijdsperiode, [Gesloten Changes], [Gesloten Incidents], [Gesloten Problems], [Gesloten Service Requests], [Instroom Changes], [Instroom Incidents], [Instroom Problems], [Instroom Service Requests])" +
                "VALUES ('"+bd+"/"+ed+"', " + data[0] + ", " + data[1] + ", " + data[2] + ", "+data[3]+ ", "+data[4]+ ", "+data[5]+ ", "+data[6]+ ", "+data[7]+")";
            SqlCommand cmd_two = new SqlCommand(query, conn);
            cmd_two.ExecuteNonQuery();
            conn.Close();
            MessageBox.Show("data is stored in dbo.data");
        }

        private void select_file_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "select file";
            openFileDialog.Filter = "excel|*.xlsx";
            openFileDialog.Multiselect = false;
            openFileDialog.ShowDialog();
            iPath.Text = openFileDialog.FileName;
        }

        private void create_sql_table_Click(object sender, EventArgs e)
        {
            string query = "CREATE TABLE [dbo].[helpDesk] (\r\n    [id]                    INT           IDENTITY (1, 1) NOT NULL,\r\n    [Bedrijf]               VARCHAR (100) NULL,\r\n    [Title]                 TEXT          NULL,\r\n    [Naam verzoeker]        VARCHAR (50)  NULL,\r\n    [Urgentie]              VARCHAR (50)  NULL,\r\n    [Tijd indienen verzoek] DATETIME2 (7) NULL,\r\n    [Tijd sluiting]         DATETIME2 (7) NULL,\r\n    [Status]                VARCHAR (50)  NULL,\r\n    [Categorie]             NVARCHAR (50) NULL,\r\n    [Subcategorie]          VARCHAR (50)  NULL,\r\n    [Type serviceverzoek]   VARCHAR (50)  NULL,\r\n    [Time to Respond]       VARCHAR (50)  NULL,\r\n    [Time to Repair]        VARCHAR (50)  NULL,\r\n    [Total Activities time] VARCHAR (50)  NULL\r\n);\r\n\r\n" +
                "CREATE TABLE [dbo].[data] (\r\n\t[id]                    INT           IDENTITY (1, 1) NOT NULL,\r\n    [Tijdsperiode]              VARCHAR (50) NULL,\r\n    [Gesloten Changes]          INT          NULL,\r\n    [Gesloten Incidents]        INT          NULL,\r\n    [Gesloten Problems]         INT          NULL,\r\n    [Gesloten Service Requests] INT          NULL,\r\n    [Instroom Changes]          INT          NULL,\r\n    [Instroom Incidents]        INT          NULL,\r\n    [Instroom Problems]         INT          NULL,\r\n    [Instroom Service Requests] INT          NULL\r\n);\r\n\r\n";
            SqlConnection conn = new SqlConnection(connectionString);
            conn.Open();
            SqlCommand cmd = new SqlCommand(query, conn);
            cmd.ExecuteNonQuery();
            MessageBox.Show("Create successfully");
            conn.Close();  
        }
    }
}
