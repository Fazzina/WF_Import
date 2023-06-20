using System;
using System.Data;
using System.Linq;
using System.Web.UI;
using ClosedXML.Excel;
using Newtonsoft.Json;
using System.Data.SqlClient;
using System.Web.UI.WebControls;
using System.Configuration;
using System.IO;
using System.Text;
using System.Web.UI.HtmlControls;
using System.Collections.Generic;

namespace WF_Import
{
    public partial class _Default : Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        public int checkEmptyTable()
        {
            string sql = "SELECT COUNT(*) from Progetti";
            using (SqlConnection conn = new SqlConnection(connectionString()))
            {
                SqlCommand cmd = new SqlCommand(sql, conn);

                try
                {
                    conn.Open();
                    int result = int.Parse(cmd.ExecuteScalar().ToString());
                    return result; // if result equals zero, then the table is empty
                }
                catch (Exception)
                {
                    throw;
                }
            }
        }

        public string connectionString()
        {
            return ConfigurationManager.ConnectionStrings["connectionString"].ConnectionString;
        }

        protected void btnUpload_Click(object sender, EventArgs e)
        {
            // Lettura del file di configurazione
            string configPath = Server.MapPath("~/ConfigExcel.json");// @"C:\Users\leonfab\source\repos\WF_Import\WF_Import\ConfigExcel.json";
            string configJson = File.ReadAllText(configPath);
            dynamic config = JsonConvert.DeserializeObject<Config>(configJson);

            // Leggi il file Excel
            var wb = new XLWorkbook(fileUpload.PostedFile.InputStream);
            var ws = wb.Worksheet(1);

            // seleziona il range di celle dalla prima all'utlima colonna utilizzata
            var range = ws.Range(ws.FirstCellUsed(), ws.LastCellUsed());

            // Crea una DataTable per contenere i dati del foglio Excel
            DataTable dt = new DataTable();

            // aggiunge colonne al DataTable in base alle intestazioni di colonna del range di celle
            foreach (var cell in range.FirstRow().Cells())
            {
                dt.Columns.Add(cell.Value.ToString());
            }

            // Aggiungi le righe alla DataTable
            int rowIndex = config.FirstRow;
            List<string> errors = new List<string>(); // lista degli errori
            foreach (var row in ws.RowsUsed().Skip(rowIndex - 1))
            {
                bool isValid = true;
                StringBuilder errorMessages = new StringBuilder();

                // Loop attraverso le colonne definite nel file di configurazione
                DataRow newRow = dt.NewRow();
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    var cellValue = row.Cell(i + 1).Value;

                    // Se il valore della cella è vuoto, aggiungi un messaggio di errore
                    if (cellValue == null || cellValue == DBNull.Value || cellValue.ToString() == "")
                    {
                        switch (config.fieldsMap[i].datatype)
                        {
                            case "date":
                                newRow[i] = DBNull.Value;
                                break;
                            default:
                                break;
                        }
                        isValid = false;
                        string columnName = dt.Columns[i].ColumnName;
                        errorMessages.AppendFormat("<img style='width: 14px;' src='x.png'/>&nbsp; On line {0} of the Excel file, the field '{1}' is blank.<br/>", row.RowNumber(), columnName);
                    }
                    else
                    {
                        // Se il valore della cella non corrisponde al tipo di dato definito nel file di configurazione, aggiungi un messaggio di errore
                        try
                        {
                            switch (config.fieldsMap[i].datatype)
                            {
                                case "int":
                                    newRow[i] = int.Parse(cellValue.ToString());
                                    break;
                                case "decimal":
                                    newRow[i] = decimal.Parse(cellValue.ToString());
                                    break;
                                case "date":
                                    newRow[i] = DateTime.Parse(cellValue.ToString());
                                    break;
                                case "string":
                                    newRow[i] = cellValue.ToString();
                                    break;
                                default:
                                    break;
                            }
                        }
                        catch (Exception)
                        {
                            isValid = false;
                            //dt.Rows.RemoveAt(i) ;
                            newRow[i] = DBNull.Value;
                            string columnName = dt.Columns[i].ColumnName;
                            errorMessages.AppendFormat("<img src='x.png'/>&nbsp; On line {0} of the Excel file, the value '{1}' of field '{2}' is invalid.<br/>", row.RowNumber(), cellValue.ToString(), columnName);
                        }
                    }
                }

                // Se la riga è valida, aggiungila alla DataTable
                if (isValid)
                {
                    dt.Rows.Add(newRow);
                }
                // Altrimenti, aggiungi un messaggio di errore alla lista
                else
                {
                    errors.Add(errorMessages.ToString());
                }
            }

            // Se ci sono errori, mostra i messaggi sulla pagina web utilizzando un controllo Label
            if (errors.Count > 0)
            {
                Label lblError = new Label();
                lblError.Text = string.Join("<br/>", errors);
                errorDiv.Controls.Add(lblError);
                errorDiv.Visible = true;

                // Aggiunge il testo al popup
                popupTextErr.InnerText = errors.Count + " errors were found.";
                popupImg.Attributes["src"] = ResolveUrl("~/warning1.png");
            }

            // Salva la DataTable in un database
            using (SqlConnection connection = new SqlConnection(connectionString()))
            {
                connection.Open();

                using (SqlBulkCopy bulkCopy = new SqlBulkCopy(connection))
                {
                    bulkCopy.DestinationTableName = "Progetti";

                    // mappatura delle colonne del DataTable alle colonne corrispondenti nella tabella del database
                    bulkCopy.ColumnMappings.Add("ANNO", "Anno");
                    bulkCopy.ColumnMappings.Add("NR. PROTOCOLLO", "nProtocollo");
                    bulkCopy.ColumnMappings.Add("DATA INSERIMENTO", "dataInserimento");
                    bulkCopy.ColumnMappings.Add("TIPOLOGIA", "Tipologia");
                    bulkCopy.ColumnMappings.Add("STATO", "Stato");
                    bulkCopy.ColumnMappings.Add("AMBITO D'INTERVENTO", "Ambito");
                    bulkCopy.ColumnMappings.Add("SOGGETTI DESTINATARI\n", "Soggetti");
                    bulkCopy.ColumnMappings.Add("TITOLO INIZIATIVA ", "Titolo");
                    bulkCopy.ColumnMappings.Add("DATA INIZIO\n(GG/MM/AA)", "dataInizio");
                    bulkCopy.ColumnMappings.Add("DATA FINE\n(GG/MM/AA)", "dataFine");

                    bulkCopy.WriteToServer(dt);
                }
            }


            // mostra il popup
            ScriptManager.RegisterStartupScript(this, GetType(), "Popup", "showPopup();", true);

            popupTextSucc.InnerText = "Operation successfully completed!";

            int counter = dt.Rows.Count;

            if (counter == 1)
            {
                popupTextRows.InnerText = counter + " row affected";
            }
            else
            {
                popupTextRows.InnerText = counter + " rows affected";
            }           

        }

        protected void redView(object sender, EventArgs e)
        {
            Response.Redirect("View.aspx", true);
        }


    }//close the partial class

}