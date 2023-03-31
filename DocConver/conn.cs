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
using Org.BouncyCastle.X509;

namespace DocConver
{
    public partial class conn : Form
    {
        private string connectionString = ConfigurationManager.ConnectionStrings["localhost"].ConnectionString;
        public conn()
        {
            InitializeComponent();
        }

        private void save(object sender, EventArgs e) //Connection String opslaan in app.config
        {
            System.Configuration.Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            config.ConnectionStrings.ConnectionStrings["localhost"].ConnectionString = cs.Text;
            config.Save(ConfigurationSaveMode.Modified);
            Application.Restart();
            Environment.Exit(0);
        }

        private void createTable(object sender, EventArgs e) //Drie sql tabllen initialiseren
        {
            try
            {
                using(SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();
                    string[] Querys = {
                        "CREATE TABLE [dbo].[helpDesk] (\r\n    [id]                    INT           IDENTITY (1, 1) NOT NULL,\r\n    [#]                     VARCHAR (50)  NULL,\r\n    [Bedrijf]               VARCHAR (MAX) NULL,\r\n    [Title]                 TEXT          NULL,\r\n    [Naam verzoeker]        TEXT          NULL,\r\n    [Urgentie]              TEXT          NULL,\r\n    [Tijd indienen verzoek] DATETIME2 (7) NULL,\r\n    [Tijd sluiting]         DATETIME2 (7) NULL,\r\n    [Status]                VARCHAR (MAX) NULL,\r\n    [Categorie]             TEXT          NULL,\r\n    [Subcategorie]          TEXT          NULL,\r\n    [Type serviceverzoek]   VARCHAR (MAX) NULL,\r\n    [Time to Respond]       TEXT          NULL,\r\n    [Time to Repair]        TEXT          NULL,\r\n    [Total Activities time] TEXT          NULL\r\n);\r\n\r\n",
                        "CREATE TABLE [dbo].[prtg] (\r\n    [id]                    INT           IDENTITY (1, 1) NOT NULL,\r\n    [Apparaat]              TEXT          NULL,\r\n    [Beschikbaarheid Stats] VARCHAR (50)  NULL,\r\n    [Percentage]            FLOAT (53)    NULL,\r\n    [Datum]                 DATETIME2 (7) NULL\r\n);\r\n",
                        "CREATE TABLE [dbo].[samenvatting] (\r\n    [id]           INT           IDENTITY (1, 1) NOT NULL,\r\n    [bedrijf]      VARCHAR (MAX) NULL,\r\n    [samenvatting] TEXT          NULL,\r\n    [time]         DATETIME      DEFAULT (getdate()) NOT NULL\r\n);\r\n\r\n"
                    };
                    for (int i = 0; i < Querys.Length; i++)
                    {
                        SqlCommand cmd = new SqlCommand(Querys[i], connection);
                        cmd.ExecuteNonQuery();
                    }
                    connection.Close();
                    MessageBox.Show("table has been create");
                }
            } catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}
