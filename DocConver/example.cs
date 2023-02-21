using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DocConver
{
    internal class example
    {
        //private void insert_to_data(object sender, EventArgs e)
        //{
        //    string bd = beginDatum.Value.ToString("yyyy-MM-dd");
        //    string ed = eindDatum.Value.ToString("yyyy-MM-dd");

        //    SqlConnection conn = new SqlConnection(connectionString);
        //    conn.Open();
        //    string[] qeurys = {
        //        "SELECT COUNT(*) FROM helpDesk WHERE [Type serviceverzoek]='Change' AND Status='Gesloten' AND [Tijd sluiting] BETWEEN '"+bd+"' AND '"+ed+"'",
        //        "SELECT COUNT(*) FROM helpDesk WHERE [Type serviceverzoek]='Incident' AND Status='Gesloten' AND [Tijd sluiting] BETWEEN '"+bd+"' AND '"+ed+"'",
        //        "SELECT COUNT(*) FROM helpDesk WHERE [Type serviceverzoek]='Problem' AND Status='Gesloten' AND [Tijd sluiting] BETWEEN '"+bd+"' AND '"+ed+"'",
        //        "SELECT COUNT(*) FROM helpDesk WHERE [Type serviceverzoek]='Request' AND Status='Gesloten' AND [Tijd sluiting] BETWEEN '"+bd+"' AND '"+ed+"'",
        //        "SELECT COUNT(*) FROM helpDesk WHERE [Type serviceverzoek]='Change' AND Status='Pending' AND [Tijd indienen verzoek] BETWEEN '"+bd+"' AND '"+ed+"'",
        //        "SELECT COUNT(*) FROM helpDesk WHERE [Type serviceverzoek]='Incident' AND Status='Pending' AND [Tijd indienen verzoek] BETWEEN '"+bd+"' AND '"+ed+"'",
        //        "SELECT COUNT(*) FROM helpDesk WHERE [Type serviceverzoek]='Problem' AND Status='Pending' AND [Tijd indienen verzoek] BETWEEN '"+bd+"' AND '"+ed+"'",
        //        "SELECT COUNT(*) FROM helpDesk WHERE [Type serviceverzoek]='Request' AND Status='Pending' AND [Tijd indienen verzoek] BETWEEN '"+bd+"' AND '"+ed+"'"
        //    };

        //    int[] data = new int[8];
        //    for (int i = 0; i < qeurys.Length; i++)
        //    {
        //        SqlCommand cmd = new SqlCommand(qeurys[i], conn);
        //        Int32 count = (Int32)(cmd.ExecuteScalar());
        //        Debug.WriteLine(qeurys[i]);
        //        data[i] = count;
        //    }
        //    string query = "INSERT INTO data (Tijdsperiode, [Gesloten Changes], [Gesloten Incidents], [Gesloten Problems], [Gesloten Service Requests], [Instroom Changes], [Instroom Incidents], [Instroom Problems], [Instroom Service Requests])" +
        //        "VALUES ('" + bd + "/" + ed + "', " + data[0] + ", " + data[1] + ", " + data[2] + ", " + data[3] + ", " + data[4] + ", " + data[5] + ", " + data[6] + ", " + data[7] + ")";
        //    SqlCommand cmd_two = new SqlCommand(query, conn);
        //    cmd_two.ExecuteNonQuery();
        //    conn.Close();
        //    MessageBox.Show("data is stored in dbo.data");
        //}
    }
}
