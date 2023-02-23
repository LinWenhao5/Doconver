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
            string query = "CREATE TABLE [dbo].[helpDesk] (\r\n    [id]                    INT           IDENTITY (1, 1) NOT NULL,\r\n    [Bedrijf]               VARCHAR (100) NULL,\r\n    [Title]                 TEXT          NULL,\r\n    [Naam verzoeker]        VARCHAR (50)  NULL,\r\n    [Urgentie]              VARCHAR (50)  NULL,\r\n    [Tijd indienen verzoek] DATETIME2 (7) NULL,\r\n    [Tijd sluiting]         DATETIME2 (7) NULL,\r\n    [Status]                VARCHAR (50)  NULL,\r\n    [Categorie]             NVARCHAR (50) NULL,\r\n    [Subcategorie]          VARCHAR (50)  NULL,\r\n    [Type serviceverzoek]   VARCHAR (50)  NULL,\r\n    [Time to Respond]       VARCHAR (50)  NULL,\r\n    [Time to Repair]        VARCHAR (50)  NULL,\r\n    [Total Activities time] VARCHAR (50)  NULL\r\n);\r\n\r\n" +
                "CREATE TABLE [dbo].[data] (\r\n\t[id]                    INT           IDENTITY (1, 1) NOT NULL,\r\n    [Tijdsperiode]              VARCHAR (50) NULL,\r\n    [Gesloten Changes]          INT          NULL,\r\n    [Gesloten Incidents]        INT          NULL,\r\n    [Gesloten Problems]         INT          NULL,\r\n    [Gesloten Service Requests] INT          NULL,\r\n    [Instroom Changes]          INT          NULL,\r\n    [Instroom Incidents]        INT          NULL,\r\n    [Instroom Problems]         INT          NULL,\r\n    [Instroom Service Requests] INT          NULL\r\n);\r\n\r\n" +
                "CREATE TABLE [dbo].[prtg] (\r\n    [id]                    INT          IDENTITY (1, 1) NOT NULL,\r\n    [Apparaat]              TEXT NULL,\r\n    [Beschikbaarheid Stats] VARCHAR (50) NULL,\r\n    [Percentage]            VARCHAR (50) NULL\r\n);\r\n\r\n";
            SqlConnection conn = new SqlConnection(connectionString);
            conn.Open();
            SqlCommand cmd = new SqlCommand(query, conn);
            cmd.ExecuteNonQuery();
            MessageBox.Show("Create successfully");
            conn.Close();
        }
    }
}
