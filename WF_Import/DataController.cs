using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using Newtonsoft.Json;

namespace WF_Import.Controllers
{
    public class DataController : ApiController
    {
        public HttpResponseMessage GetProjects()
        {
            try
            {
                // Connessione al database e recupero dei dati
                var connectionString = ConfigurationManager.ConnectionStrings["connectionString"].ConnectionString;
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
                    var json = JsonConvert.SerializeObject(rows, jsonSettings);

                    //var projectsJson = JsonConvert.SerializeObject(projects);
                    var response = new HttpResponseMessage(HttpStatusCode.OK);
                    response.Content = new StringContent(json, System.Text.Encoding.UTF8, "application/json");
                    return response;
                }
            }
            catch (Exception ex)
            {
                var response = new HttpResponseMessage(HttpStatusCode.InternalServerError);
                response.Content = new StringContent(ex.Message);
                return response;
            }
        }

    }
}