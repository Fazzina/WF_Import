using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;

namespace WF_Import
{
    public partial class MyEndpoint : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            // Connessione al database e recupero dei dati
            var connectionString = "Data Source=VT-LEONFAB;Initial Catalog=ProgettiSocialiCustom;User ID=sa;Password=6LGmG!iDrzq4";
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                string query = "SELECT * FROM Progetti order by nProtocollo";
                SqlCommand command = new SqlCommand(query, connection);
                SqlDataReader reader = command.ExecuteReader();

                // Trasformazione dei dati in JSON
                List<Dictionary<string, object>> rows = new List<Dictionary<string, object>>();
                while (reader.Read())
                {
                    Dictionary<string, object> row = new Dictionary<string, object>();
                    for (int i = 0; i < reader.FieldCount; i++)
                    {
                        row.Add(reader.GetName(i), reader[i]);
                    }
                    rows.Add(row);
                }

                var jsonSettings = new JsonSerializerSettings();
                jsonSettings.DateFormatString = "dd/MM/yyyy";

                //string json = new JavaScriptSerializer().Serialize(rows);
                string json = JsonConvert.SerializeObject(rows, jsonSettings);
                // Restituzione della risposta JSON
                Response.Write(json);
                Response.End();
                connection.Close();
            }
        }
        
    }
}