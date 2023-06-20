/*
using System;
using System.Data;
using System.Linq;
using System.Web.UI;
using ClosedXML.Excel;
using Newtonsoft.Json;
using System.Data.SqlClient;
using System.Collections.Generic;
using Newtonsoft.Json.Converters;
using System.Threading.Tasks;
using System.Net.Http;
using System.Web.UI.WebControls;
using System.Configuration;
using System.Globalization;

namespace WF_Import
{
    public partial class _Default : Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            projectsGridView.Visible = false;
            if (projectsGridView.Visible == true)
            {
                HideButton.Visible = true;
            }
            else
            {
                ShowButton.Visible = false;
            }
        }


        public string connectionString()
        {
            return ConfigurationManager.ConnectionStrings["connectionString"].ConnectionString;
        }


        public DataTable fromExcelToDataTable()
        {
            try
            {
                // crea un nuovo oggetto cartella di lavoro dal file caricato
                var workbook = new XLWorkbook(fileUpload.PostedFile.InputStream);

                // seleziona il primo foglio di lavoro della cartella di lavoro
                var worksheet = workbook.Worksheet(1);

                // seleziona il range di celle dalla prima all'utlima colonna utilizzata
                var range = worksheet.Range(worksheet.FirstCellUsed(), worksheet.LastCellUsed());

                // crea un nuovo DataTable
                var table = new DataTable();

                // aggiunge colonne al DataTable in base alle intestazioni di colonna del range di celle
                foreach (var cell in range.FirstRow().Cells())
                {
                    table.Columns.Add(cell.Value.ToString());
                }

                // aggiunge una colonna "Errori" al DataTable
                table.Columns.Add("Errori", typeof(string));


                // aggiunge righe al DataTable in base ai dati del range di celle
                foreach (var row in range.RowsUsed().Skip(1))
                {
                    var data = new List<object>();
                    bool hasEmptyCell = false;
                    bool hasWrongType = false;
                    string columnName = string.Empty;
                    for (int i = 0; i < table.Columns.Count - 1; i++)
                    {
                        var cell = row.Cell(i + 1);
                        var expectedType = table.Columns[i].DataType;
                        object value = null;
                        if (cell.IsEmpty())
                        {
                            value = DBNull.Value;
                            columnName = table.Columns[i].ColumnName;
                            hasEmptyCell = true;
                        }
                        else if (table.Rows.GetType().Name != expectedType.Name)
                        {
                            value = DBNull.Value;
                            columnName = table.Columns[i].ColumnName;
                            hasWrongType = true;
                        }
                        else
                        {
                            // tenta di convertire il valore della cella nel tipo di dato atteso
                            value = Convert.ChangeType(cell.Value, expectedType);
                        }

                        data.Add(value);
                        
                        System.Diagnostics.Debug.WriteLine("##################################################");
                        System.Diagnostics.Debug.WriteLine(value);
                        System.Diagnostics.Debug.WriteLine("##################################################");
                        System.Diagnostics.Debug.WriteLine(data);
                        
                    }

                    if (hasEmptyCell)
                    {
                        // se la riga contiene almeno una cella vuota, segnala un errore
                        table.Rows.Add(data.Concat(new string[] { "Il campo " + columnName + " è vuoto" }).ToArray());
                    }
                    else if (hasWrongType)
                    {
                        table.Rows.Add(data.Concat(new object[] { "Il campo " + columnName + " contiene un valore non valido" }).ToArray());
                    }
                    else
                    {
                        // altrimenti aggiunge la riga al DataTable
                        table.Rows.Add(data.ToArray());
                    }
                }
                // fino a qua ci siamo recupera i dati e li sposta in una datatable

                return table;
            }
            catch (Exception)
            {
                throw;
            }

        }


        public void fromDataTabletoDatabase(DataTable t)
        {
            try
            {
                // connessione al database
                SqlConnection connection = new SqlConnection(connectionString());

                // creazione di un oggetto SqlBulkCopy per l'inserimento dei dati nel database
                SqlBulkCopy bulkCopy = new SqlBulkCopy(connection);
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
                bulkCopy.ColumnMappings.Add("Errori", "Errori");


                connection.Open();
                bulkCopy.WriteToServer(t);

                int counter = t.Rows.Count;

                if (counter == 1)
                {
                    lblRowCount.Text = "É stata interessata " + counter + " riga.";
                }
                else
                {
                    lblRowCount.Text = "Sono state interessate " + counter + " righe.";
                }
                connection.Close();
            }
            catch (Exception)
            {
                throw;
            }
        }

        public async void JsonToGridView()
        {
            // recupera i dati dal database tramite JSON
            var projectsJson = await GetProjectsJsonFromAPI();

            // deserializza i dati JSON in una lista di oggetti
            var projects = JsonConvert.DeserializeObject<List<Progetto>>(projectsJson, new JsonSerializerSettings
            {
                Error = (sender2, args) =>
                {
                    if (args.ErrorContext.Error.GetType() == typeof(JsonReaderException))
                    {
                        args.ErrorContext.Handled = true;
                    }
                },
                Converters = { new IsoDateTimeConverter { DateTimeFormat = "dd/MM/yyyy" } }
            });

            if (projects.Count != 0)
            {
                // mostra i dati
                panelData.Visible = true;

                // assegna la lista di oggetti come sorgente dati del GridView
                projectsGridView.DataSource = projects;
                projectsGridView.DataBind();
            }

        }

        private async Task<string> GetProjectsJsonFromAPI()
        {
            using (var client = new HttpClient())
            {
                var response = await client.GetAsync("https://localhost:44359/api/data/getprojects");
                if (response.IsSuccessStatusCode)
                {
                    return await response.Content.ReadAsStringAsync();
                }
                else
                {
                    throw new Exception($"Error getting projects: {response.StatusCode}");
                }
            }
        }

        protected void btnUpload_Click(object sender, EventArgs e)
        {
            try
            {
                if (fileUpload.HasFile)
                {
                    fromDataTabletoDatabase(fromExcelToDataTable());

                    JsonToGridView();

                    // nasconde lo spinner
                    //ClientScript.RegisterStartupScript(GetType(), "hideSpinner", "<script>hideSpinner();</script>");

                    // visualizza il GridView
                    projectsGridView.Visible = true;
                    //HideButton.Visible = true;
                    //ShowButton.Visible = false;
                }
                else
                {
                    //ClientScript.RegisterStartupScript(GetType(), "showAlert", "<script>showAlert();</script>");
                    projectsGridView.Visible = true;
                    //HideButton.Visible = true;
                    //ShowButton.Visible = false;
                }
            }
            catch (Exception)
            {

                throw;
            }
        }

        protected void btnReset_Click(object sender, EventArgs e)
        {
            // connessione al database
            using (SqlConnection connection = new SqlConnection(connectionString()))
            {
                connection.Open();
                using (SqlCommand command = new SqlCommand("DELETE FROM Progetti", connection))
                {
                    command.ExecuteNonQuery();
                }
            }

            panelData.Visible = false;
            lblRowCount.Visible = false;
        }

        protected void btnShow_Click(object sender, EventArgs e)
        {
            panelData.Visible = true;
            projectsGridView.Visible = true;
            HideButton.Visible = true;
            if (HideButton.Visible == true)
            {
                ShowButton.Visible = false;
            }
        }
        protected void btnHide_Click(object sender, EventArgs e)
        {
            panelData.Visible = false;
            ShowButton.Visible = true;
            if (ShowButton.Visible == true)
            {
                HideButton.Visible = false;
            }
        }

        protected void convertErrorRowColor(object sender, GridViewRowEventArgs e)
        {
            if (e.Row.RowType == DataControlRowType.DataRow)
            {
                string errori = DataBinder.Eval(e.Row.DataItem, "Errori")?.ToString();
                if (!string.IsNullOrEmpty(errori))
                {
                    e.Row.CssClass = "alert alert-danger";
                    e.Row.Font.Bold = true;
                }
            }
        }

    }//close the partial class

}
<asp:LinkButton ID="ImportButton" runat="server" tagname="ImportButt" OnClick="btnUpload_Click" OnClientClick="showSpinner();" CssClass="btn btn-primary"><i class="bi bi-upload"></i>&nbsp;Import</asp:LinkButton>

<!-- OnRowDataBound="convertErrorRowColor" -->

<asp:LinkButton ID="ResetButton" CausesValidation="False" runat="server" OnClick="btnReset_Click" CssClass="btn btn-secondary"><i class="bi bi-trash"></i>&nbsp;Reset</asp:LinkButton>
        <asp:LinkButton ID="HideButton" CausesValidation="False" runat="server" OnClick="btnHide_Click" CssClass="btn btn-secondary"><i class="bi bi-eye-slash"></i>&nbsp;Hide</asp:LinkButton>
    <asp:LinkButton ID="ShowButton" CausesValidation="False" runat="server" OnClick="btnShow_Click" CssClass="btn btn-secondary"><i class="bi bi-eye"></i>&nbsp;Show</asp:LinkButton>




    {
  "file_structure": {
    "headers_row": 1,
    "data_start_row": 2,
    "columns": [
      {
        "columnID": 1,
        "datatype": "string",
        "required": true
      },
      {
        "columnID": 2,
        "datatype": "int",
        "required": true
      },
      {
        "columnID": 3,
        "datatype": "date",
        "required": true
      },
      {
        "columnID": 4,
        "datatype": "string",
        "required": true
      },
      {
        "columnID": 5,
        "datatype": "string",
        "required": true
      },
      {
        "columnID": 6,
        "datatype": "string",
        "required": true
      },
      {
        "columnID": 7,
        "datatype": "string",
        "required": true
      },
      {
        "columnID": 8,
        "datatype": "string",
        "required": true
      },
      {
        "columnID": 9,
        "datatype": "date",
        "required": true
      },
      {
        "columnID": 10,
        "datatype": "date",
        "required": true
      }
    ]
  }
}
*/

/*
###################################################################################

using System;
using System.Data;
using System.Linq;
using System.Web.UI;
using ClosedXML.Excel;
using Newtonsoft.Json;
using System.Data.SqlClient;
using System.Collections.Generic;
using Newtonsoft.Json.Converters;
using System.Threading.Tasks;
using System.Net.Http;
using System.Web.UI.WebControls;
using System.Configuration;
using System.Globalization;
using System.IO;
using System.Text;

namespace WF_Import
{
    public partial class _Default : Page
    {

        protected void Page_Load(object sender, EventArgs e)
        {
            projectsGridView.Visible = false;
            
            if (projectsGridView.Visible == true)
            {
                HideButton.Visible = true;
            }
            else
            {
                ShowButton.Visible = false;
            }
            
        }


        public string connectionString()
{
    return ConfigurationManager.ConnectionStrings["connectionString"].ConnectionString;
}

protected void btnUpload_Click(object sender, EventArgs e)
{
    // Lettura del file di configurazione
    string configPath = @"C:\Users\leonfab\source\repos\WF_Import\WF_Import\ConfigExcel.json";
    string configJson = File.ReadAllText(configPath);
    Config config = JsonConvert.DeserializeObject<Config>(configJson);

    // Leggi il file Excel
    var workbook = new XLWorkbook(fileUpload.PostedFile.InputStream);
    var worksheet = workbook.Worksheet(1);

    // Ottiene l'ultima riga del foglio di lavoro
    int lastRow = worksheet.LastRowUsed().RowNumber();

    // seleziona il range di celle dalla prima all'utlima colonna utilizzata
    var range = worksheet.Range(worksheet.FirstCellUsed(), worksheet.LastCellUsed());

    // Imposta una DataTable per contenere i dati
    DataTable dataTable = new DataTable();

    // aggiunge colonne al DataTable in base alle intestazioni di colonna del range di celle
    foreach (var cell in range.FirstRow().Cells())
    {
        dataTable.Columns.Add(cell.Value.ToString());
    }

    // aggiunge una colonna "Errori" al DataTable
    dataTable.Columns.Add("ERRORI", typeof(string));

    // Itera su ogni riga del foglio di lavoro, a partire dalla FirstRow definita nel file di configurazione
    for (int row = config.FirstRow; row <= lastRow; row++)
    {
        DataRow dataRow = dataTable.NewRow();
        bool rowIsValid = true;

        // Itera su ogni colonna della riga, come definito nel file di configurazione
        foreach (FieldMap fieldMap in config.fieldsMap)
        {
            int columnID = fieldMap.columnID;
            string cellValue = worksheet.Cell(row, columnID).GetString();

            // Convalida i dati in base al tipo di dato previsto
            if (string.IsNullOrEmpty(cellValue) && !fieldMap.allowNull)
            {
                dataRow["ERRORI"] = "Il campo " + fieldMap.columnName + " è vuoto";
                rowIsValid = false;
            }
            else
            {
                try
                {
                    dataRow[fieldMap.columnName] = Convert.ChangeType(cellValue, GetDataType(fieldMap.datatype));
                }
                catch (Exception)
                {
                    dataRow["ERRORI"] = "Il campo " + fieldMap.columnName + " non è del tipo previsto";
                    rowIsValid = false;
                }
            }
        }

        // Aggiunge la riga alla DataTable se è valida
        if (rowIsValid)
        {
            dataTable.Rows.Add(dataRow);
        }


        // Esporta i dati nel database utilizzando SqlBulkCopy
        using (SqlConnection connection = new SqlConnection(connectionString()))
        {
            connection.Open();

            using (SqlBulkCopy bulkCopy = new SqlBulkCopy(connection))
            {
                bulkCopy.DestinationTableName = "Progetti";

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
                bulkCopy.ColumnMappings.Add("ERRORI", "Errori");

                bulkCopy.WriteToServer(dataTable);
            }
        }
    }

    // assegna la lista di oggetti come sorgente dati del GridView
    projectsGridView.DataSource = dataTable;
    projectsGridView.DataBind();
    projectsGridView.Visible = true;
}

private static Type GetDataType(string datatype)
{
    switch (datatype.ToLower())
    {
        case "string":
            return typeof(string);
        case "int":
            return typeof(int);
        case "date":
            return typeof(DateTime);
        default:
            throw new ArgumentException("Tipo di dati non valido: " + datatype);
    }
}

        

    }//close the partial class

}
*/

/*
 <Columns>
                    <asp:BoundField DataField="Anno" HeaderText="Anno" />
                    <asp:BoundField DataField="nProtocollo" HeaderText="Nr. Protocollo" />
                    <asp:BoundField DataField="dataInserimento" HeaderText="Data Inserimento&#10;(GG/MM/AA)" DataFormatString="{0:dd/MM/yyyy}" />
                    <asp:BoundField DataField="Tipologia" HeaderText="Tipologia" />
                    <asp:BoundField DataField="Stato" HeaderText="Stato" />
                    <asp:BoundField DataField="Ambito" HeaderText="Ambito d'intervento" />
                    <asp:BoundField DataField="Soggetti" HeaderText="Soggetti Destinatari" />
                    <asp:BoundField DataField="Titolo" HeaderText="Titolo Iniziativa" />
                    <asp:BoundField DataField="dataInizio" HeaderText="Data Inizio&#10;(GG/MM/AA)" DataFormatString="{0:dd/MM/yyyy}" />
                    <asp:BoundField DataField="dataFine" HeaderText="Data Fine&#10;(GG/MM/AA)" DataFormatString="{0:dd/MM/yyyy}" />
                    <asp:BoundField DataField="Errori" HeaderText="Errori" />
                </Columns>
 */


/*
 ##############################################################################################
QUESTO FUNZIONA TUTTO

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

namespace WF_Import
{
    public partial class _Default : Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {


            projectsGridView.Visible = false;

            if (checkEmptyTable() == 0)
            {
                ShowButton.Visible = false;
            }
            else
            {
                ShowButton.Visible = true;
            }

            #####
             if (projectsGridView.Visible == true)
                {
                    HideButton.Visible = true;
                }
                else
                {
                    ShowButton.Visible = false;
                }
             #####
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

    // aggiunge una colonna "Errori" al DataTable
    dt.Columns.Add("ERRORI", typeof(string));

    // Aggiungi le righe alla DataTable
    int rowIndex = config.FirstRow;
    foreach (var row in ws.RowsUsed().Skip(rowIndex - 1))
    {
        DataRow newRow = dt.NewRow();
        for (int i = 0; i < dt.Columns.Count; i++)
        {
            newRow[i] = row.Cell(i + 1).Value;
        }
        dt.Rows.Add(newRow);
    }



    // Loop attraverso le righe della DataTable
    foreach (DataRow row in dt.Rows)
    {
        bool isValid = true;
        StringBuilder errorMessages = new StringBuilder();
        string columnName = string.Empty;
        // Loop attraverso le colonne definite nel file di configurazione
        foreach (var field in ((Config)config).fieldsMap)
        {
            var cellValue = row[field.columnID - 1];

            // Se il valore della cella è vuoto, aggiungi un messaggio di errore
            if (cellValue == null || cellValue == DBNull.Value || cellValue.ToString() == "")
            {
                switch (field.datatype)
                {
                    case "date":
                        row[field.columnID - 1] = DBNull.Value;
                        break;
                    default:
                        break;
                }
                isValid = false;
                columnName = dt.Columns[field.columnID - 1].ColumnName;
                errorMessages.AppendFormat("Il campo '{0}' è vuoto.\n", columnName);
            }
            else
            {
                // Se il valore della cella non corrisponde al tipo di dato definito nel file di configurazione, aggiungi un messaggio di errore
                try
                {
                    switch (field.datatype)
                    {
                        case "int":
                            int.Parse(cellValue.ToString());
                            break;
                        case "decimal":
                            decimal.Parse(cellValue.ToString());
                            break;
                        case "date":
                            DateTime.Parse(cellValue.ToString());
                            break;
                        case "string":
                            break;
                        default:
                            break;
                    }
                }
                catch (Exception)
                {
                    isValid = false;
                    row[field.columnID - 1] = DBNull.Value;
                    columnName = dt.Columns[field.columnID - 1].ColumnName;
                    errorMessages.AppendFormat("Il valore '{0}' del campo '{1}' non è valido.\n", cellValue.ToString(), columnName);
                }
            }
        }

        // Se la riga non è valida, aggiungi un messaggio di errore nella colonna "Errori"
        if (!isValid)
        {
            row["Errori"] = errorMessages.ToString();
        }
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
            bulkCopy.ColumnMappings.Add("ERRORI", "Errori");

            bulkCopy.WriteToServer(dt);
        }
    }

    int counter = dt.Rows.Count;

    if (counter == 1)
    {
        lblRowCount.Text = "É stata interessata " + counter + " riga.";
    }
    else
    {
        lblRowCount.Text = "Sono state interessate " + counter + " righe.";
    }

    panelData.Visible = true;

    projectsGridView.DataBind();

    projectsGridView.Visible = true;


}

protected void btnReset_Click(object sender, EventArgs e)
{
    // connessione al database
    using (SqlConnection connection = new SqlConnection(connectionString()))
    {
        connection.Open();
        using (SqlCommand command = new SqlCommand("DELETE FROM Progetti", connection))
        {
            command.ExecuteNonQuery();
        }
    }

    if (checkEmptyTable() == 0)
    {
        ShowButton.Visible = false;
    }
    else
    {
        ShowButton.Visible = true;
    }

    panelData.Visible = false;
    lblRowCount.Visible = false;
}

protected void btnShow_Click(object sender, EventArgs e)
{
    panelData.Visible = true;
    projectsGridView.Visible = true;
    HideButton.Visible = true;
    if (HideButton.Visible == true)
    {
        ShowButton.Visible = false;
    }
}
protected void btnHide_Click(object sender, EventArgs e)
{
    panelData.Visible = false;
    ShowButton.Visible = true;
    if (ShowButton.Visible == true)
    {
        HideButton.Visible = false;
    }
}

protected void convertErrorRowColor(object sender, GridViewRowEventArgs e)
{
    if (e.Row.RowType == DataControlRowType.DataRow)
    {
        string errori = DataBinder.Eval(e.Row.DataItem, "Errori")?.ToString();
        if (!string.IsNullOrEmpty(errori))
        {
            e.Row.CssClass = "alert alert-danger";
            e.Row.Font.Bold = true;
        }
    }
}


protected void btnExport_Click(object sender, EventArgs e)
{
    var query = $"SELECT * FROM ProgettiEx ORDER BY p";
    var dataTable = new DataTable();

    using (var connection = new SqlConnection(connectionString()))
    {
        var adapter = new SqlDataAdapter(query, connection);

        adapter.Fill(dataTable);
    }

    foreach (DataRow row in dataTable.Rows)
    {
        foreach (DataColumn column in dataTable.Columns)
        {
            var cellValue = row[column];
            if (column.DataType == typeof(DateTime))
            {
                var date = Convert.ToDateTime(cellValue);
                row[column] = date.ToString("dd/MM/yyyy");
            }
        }
    }

    var fileName = "ProgettiEx.xlsx";
    var filePath = Server.MapPath("~/") + fileName;

    using (var workbook = new XLWorkbook())
    {
        var worksheet = workbook.Worksheets.Add("ProgettiEx");
        worksheet.Cell(1, 1).InsertTable(dataTable);
        workbook.SaveAs(filePath);
    }

    // Invia il file come allegato nella risposta HTTP
    Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    Response.AppendHeader("Content-Disposition", $"attachment; filename={fileName}");
    Response.TransmitFile(filePath);
    Response.End();
    string myScriptValue = "function hideSpinner() {document.getElementById('spinner').style.display = 'none';}";
    ScriptManager.RegisterClientScriptBlock(this, GetType(), "myScriptName", myScriptValue, true);

    lblSuccess.Text = "Dati esportati con successo";
}

    }//close the partial class

}



<Columns>
                    <asp:BoundField DataField="Anno" HeaderText="Anno" />
                    <asp:BoundField DataField="nProtocollo" HeaderText="Nr. Protocollo" />
                    <asp:BoundField DataField="dataInserimento" HeaderText="Data Inserimento&#10;(GG/MM/AA)" DataFormatString="{0:dd/MM/yyyy}" />
                    <asp:BoundField DataField="Tipologia" HeaderText="Tipologia" />
                    <asp:BoundField DataField="Stato" HeaderText="Stato" />
                    <asp:BoundField DataField="Ambito" HeaderText="Ambito d'intervento" />
                    <asp:BoundField DataField="Soggetti" HeaderText="Soggetti Destinatari" />
                    <asp:BoundField DataField="Titolo" HeaderText="Titolo Iniziativa" />
                    <asp:BoundField DataField="dataInizio" HeaderText="Data Inizio&#10;(GG/MM/AA)" DataFormatString="{0:dd/MM/yyyy}" />
                    <asp:BoundField DataField="dataFine" HeaderText="Data Fine&#10;(GG/MM/AA)" DataFormatString="{0:dd/MM/yyyy}" />
                </Columns>



<!--
        
    <script type="text/javascript">
        // make an AJAX call to retrieve the data from the server in JSON format
        $.ajax({
            url: 'https://localhost:44359/api/data/getprojects',
            method: 'GET',
            dataType: 'json',
            beforeSend: function () {
                document.getElementById("spinner").style.display = "block"
            },
            success: function () {
                console.log('JSON caricati')
                document.getElementById("spinner").style.display = "none";
            }
        });
    </script>

    <script>
        document.forms[0].onsubmit = function () { showSpinner(); }
    </script>
        -->




<asp:SqlDataSource ID="SqlDataSource1" DataSourceMode="DataSet" runat="server" ConnectionString="<%$ ConnectionStrings:connectionString%>" SelectCommand="SELECT [Anno] as Anno,[nProtocollo] as [Nr. Protocollo],[dataInserimento] as [Data Inserimento],[Tipologia],[Stato],[Ambito] as [Ambito d'intervento],[Soggetti] as [Soggetti Destinatari],[Titolo] as [Titolo Iniziativa],[dataInizio] as [Data Inizio (GG/MM/AA)],[dataFine] as [Data Fine (GG/MM/AA)] FROM [Progetti] ORDER BY nProtocollo"></asp:SqlDataSource>

            if (SearchContron.Text.Length > 0)
            {
                SqlDataSource1.SelectCommand = "SELECT [Anno] as Anno,[nProtocollo] as [Nr. Protocollo],[dataInserimento] as [Data Inserimento],[Tipologia],[Stato],[Ambito] as [Ambito d'intervento],[Soggetti] as [Soggetti Destinatari],[Titolo] as [Titolo Iniziativa],[dataInizio] as [Data Inizio (GG/MM/AA)],[dataFine] as [Data Fine (GG/MM/AA)] FROM [Progetti] WHERE (nProtocollo LIKE N'%" + SearchContron.Text + "%' or Anno LIKE N'%" + SearchContron.Text + "%' or dataInserimento LIKE N'%" + SearchContron.Text + "%' or Tipologia LIKE N'%" + SearchContron.Text + "%' or Stato LIKE N'%" + SearchContron.Text + "%' or Ambito LIKE N'%" + SearchContron.Text + "%' or Soggetti LIKE N'%" + SearchContron.Text + "%' or Titolo LIKE N'%" + SearchContron.Text + "%' or dataInizio LIKE N'%" + SearchContron.Text + "%' or dataFine LIKE N'%" + SearchContron.Text + "%' ) ORDER BY nProtocollo";
            }
            else
            {
                SqlDataSource1.SelectCommand = "SELECT [Anno] as Anno,[nProtocollo] as [Nr. Protocollo],[dataInserimento] as [Data Inserimento],[Tipologia],[Stato],[Ambito] as [Ambito d'intervento],[Soggetti] as [Soggetti Destinatari],[Titolo] as [Titolo Iniziativa],[dataInizio] as [Data Inizio (GG/MM/AA)],[dataFine] as [Data Fine (GG/MM/AA)] FROM [Progetti] ORDER BY nProtocollo";
            }






<asp:CommandField ButtonType="Button" ShowEditButton="true">
                        <asp:ControlStyle BackColor="#339933" ForeColor="White" />
                    </asp:CommandField>
                    <asp:CommandField ButtonType="Button" ShowDeleteButton="true">
                        <asp:ControlStyle BackColor="Red" ForeColor="White" />
                    </asp:CommandField>


            var query = $"SELECT [Anno] as Anno,[nProtocollo] as [Nr. Protocollo],[dataInserimento] as [Data Inserimento],[Tipologia],[Stato],[Ambito] as [Ambito d'intervento],[Soggetti] as [Soggetti Destinatari],[Titolo] as [Titolo Iniziativa],[dataInizio] as [Data Inizio (GG/MM/AA)],[dataFine] as [Data Fine (GG/MM/AA)] FROM [Progetti] ORDER BY nProtocollo";
            var dataTable = new DataTable();

            using (var connection = new SqlConnection(connectionString()))
            {
                var adapter = new SqlDataAdapter(query, connection);

                adapter.Fill(dataTable);
            }

            

*/

/*
 <div class="form-group left">
                <label for="anno">Anno:</label>
                <asp:TextBox ID="anno" runat="server"></asp:TextBox>
            </div>

            <div class="form-group left">
                <label for="protocollo">Nr. Protocollo:</label>
                <asp:TextBox ID="protocollo" runat="server"></asp:TextBox>
            </div>

            <div class="form-group left">
                <label for="datainserimento">Data Inserimento (GG/MM/AA):</label>
                <asp:TextBox ID="datainserimento" runat="server"></asp:TextBox>
            </div>

            <div class="form-group left">
                <label for="Tipologia">Tipologia:</label>
                <asp:TextBox ID="tipologia" runat="server"></asp:TextBox>
            </div>

            <div class="form-group left">
                <label for="stato">Stato:</label>
                <asp:TextBox ID="stato" runat="server"></asp:TextBox>
            </div>

            <div class="form-group right">
                <label for="ambito">Ambito d'intervento:</label>
                <asp:TextBox ID="ambito" runat="server"></asp:TextBox>
            </div>

            <div class="form-group right">
                <label for="soggetti">Soggetti Destinatari:</label>
                <asp:TextBox ID="soggetti" runat="server"></asp:TextBox>
            </div>

            <div class="form-group right">
                <label for="titolo">Titolo Iniziativa:</label>
                <asp:TextBox ID="titolo" runat="server"></asp:TextBox>
            </div>

            <div class="form-group right">
                <label for="datainizio">Data Inizio (GG/MM/AA):</label>
                <asp:TextBox ID="datainizio" runat="server"></asp:TextBox>
            </div>

            <div class="form-group right">
                <label for="datafine">Data Fine (GG/MM/AA):</label>
                <asp:TextBox ID="datafine" runat="server"></asp:TextBox>
            </div>



<style>
        .form-group {
            margin-bottom: 20px;
            color: black;
        }

        label {
            display: block;
            font-weight: bold;
            margin-bottom: 5px;
        }

        input[type="text"] {
            padding: 10px;
            border-radius: 5px;
            border: 1px solid #ccc;
            width: 100%;
            font-size: 16px;
        }

        .left {
            float: left;
        }

        .right {
            float: right;
        }
</style> 
 */