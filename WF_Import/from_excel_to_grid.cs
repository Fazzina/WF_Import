using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace WF_Import
{/*
    protected void ImportExcel(object sender, EventArgs e)  //view data from an imported excel to a table in the webpage
    {
        isLoading = true;
        isLoadingHiddenField.Value = isLoading.ToString();

        //Save the uploaded Excel file.
        string filePath = Server.MapPath("") + Path.GetFileName(FileUpload1.PostedFile.FileName);
        FileUpload1.SaveAs(filePath);

        //Open the Excel file using ClosedXML.
        using (XLWorkbook workBook = new XLWorkbook(filePath))
        {
            //Read the first Sheet from Excel file.
            IXLWorksheet workSheet = workBook.Worksheet(1);

            //Create a new DataTable.
            DataTable dt = new DataTable();

            //Loop through the Worksheet rows.
            bool firstRow = true;
            foreach (IXLRow row in workSheet.Rows())
            {
                //Use the first row to add columns to DataTable.
                if (firstRow)
                {
                    foreach (IXLCell cell in row.Cells())
                    {
                        dt.Columns.Add(cell.Value.ToString());
                    }
                    firstRow = false;
                }
                else
                {
                    //Add rows to DataTable.
                    dt.Rows.Add();
                    int i = 0;
                    foreach (IXLCell cell in row.Cells())
                    {
                        dt.Rows[dt.Rows.Count - 1][i] = cell.Value.ToString();
                        i++;
                    }
                }

                GridView1.DataSource = dt;
                GridView1.DataBind();

                isLoading = false;
                isLoadingHiddenField.Value = isLoading.ToString();
            }
        }
    }*/



    //vari tentativi
    /*
        protected void btnUpload_Click(object sender, EventArgs e)
        {
            // Step 2: Read Excel file and store in DataTable
            var workbook = new XLWorkbook(fileUpload.FileContent);
            var worksheet = workbook.Worksheet(1);
            var dataTable = worksheet.RangeUsed().AsDataTable();

            // Step 3: Create SQL Server table
            var connectionString = "Data Source=MyServer;Initial Catalog=MyDatabase;Integrated Security=True;";
            var tableName = "MyTable";
            var createTableSql = $"CREATE TABLE {tableName} ({string.Join(", ", dataTable.Columns.Cast<DataColumn>().Select(c => $"{c.ColumnName} NVARCHAR(MAX)"))})";
            using (var connection = new SqlConnection(connectionString))
            {
                connection.Open();
                using (var command = new SqlCommand(createTableSql, connection))
                {
                    command.ExecuteNonQuery();
                }
            }

            // Step 4: Insert data into SQL Server table
            var insertSql = $"INSERT INTO {tableName} ({string.Join(", ", dataTable.Columns.Cast<DataColumn>().Select(c => c.ColumnName))}) VALUES ({string.Join(", ", dataTable.Columns.Cast<DataColumn>().Select(c => $"@{c.ColumnName}"))})";
            using (var connection = new SqlConnection(connectionString))
            {
                connection.Open();
                using (var command = new SqlCommand(insertSql, connection))
                {
                    foreach (DataRow row in dataTable.Rows)
                    {
                        foreach (DataColumn col in dataTable.Columns)
                        {
                            command.Parameters.AddWithValue($"@{col.ColumnName}", row[col]);
                        }
                        command.ExecuteNonQuery();
                        command.Parameters.Clear();
                    }
                }
            }

            // Step 5: Convert data to JSON
            var selectSql = $"SELECT * FROM {tableName}";
            var dataTableJson = JsonConvert.SerializeObject(dataTable);
            var sqlTableJson = "";
            using (var connection = new SqlConnection(connectionString))
            {
                connection.Open();
                using (var command = new SqlCommand(selectSql, connection))
                {
                    using (var reader = command.ExecuteReader())
                    {
                        var sqlTable = new DataTable();
                        sqlTable.Load(reader);
                        sqlTableJson = JsonConvert.SerializeObject(sqlTable);
                    }
                }
            }
        }
        

        protected void btnUpload_Click(object sender, EventArgs e)
        {
            // Replace the connection string with your own
            string connectionString = "Data Source=VT-LEONFAB;Initial Catalog=ProgettiSocialiCustom;User ID=sa;Password=6LGmG!iDrzq4";

            // Replace the table name and column names with your own
            string tableName = "Progetti";
            string[] columnNames = { "Anno", "nProtocollo", "dataInserimento", "Tipologia", "Stato", "Ambito", "Soggetti", "Titolo", "dataInizio", "dataFine" };

            // Set the maximum length for each SQL column
            int[] columnMaxLengths = { 4, 10, 50, 50, 50, 50, 50, 50, 50, 50 };

            // Load the Excel file using ClosedXML
            using (XLWorkbook workbook = new XLWorkbook(fileUpload.FileContent))
            {
                // Get the first worksheet in the file
                IXLWorksheet worksheet = workbook.Worksheet("MAPPATURA 2017");

                // Get the range of cells containing data in the worksheet, excluding any filtered rows or columns
                var range = worksheet.Range(
                    worksheet.FirstCellUsed().Address.RowNumber + 1, // First row with data
                    1, // First column
                    worksheet.LastCellUsed().Address.RowNumber, // Last row with data
                    worksheet.LastCellUsed().Address.ColumnNumber // Last column
                ).Clear(XLClearOptions.AllFormats); // Exclude filtered rows and columns

                // Insert a row above the header row to add column headers
                var headersRow = range.InsertRowsAbove(1).First();

                // Set the header name for each column
                for (int i = 0; i < columnNames.Length; i++)
                {
                    headersRow.Cell(i + 1).Value = columnNames[i];
                }

                // Get the data from the range as a DataTable, truncating values that exceed the maximum length
                DataTable dataTable = new DataTable();
                foreach (IXLCell cell in range.FirstRow().Cells())
                {
                    dataTable.Columns.Add(cell.GetString(), typeof(string));
                }
                foreach (IXLRangeRow row in range.Rows())
                {
                    DataRow dataRow = dataTable.NewRow();
                    for (int i = 0; i < dataTable.Columns.Count; i++)
                    {
                        string value = row.Cell(i + 1).GetString();
                        if (value.Length > columnMaxLengths[i])
                        {
                            value = value.Substring(0, columnMaxLengths[i]);
                        }
                        dataRow[i] = value;
                    }
                    dataTable.Rows.Add(dataRow);
                }

                // Insert data into SQL Server table
                // Step 4: Insert data into SQL Server table
                var insertSql = $"INSERT INTO {tableName} ({string.Join(", ", dataTable.Columns.Cast<DataColumn>().Select(c => c.ColumnName))}) VALUES ({string.Join(", ", dataTable.Columns.Cast<DataColumn>().Select(c => $"@{c.ColumnName}"))})";
                using (var connection = new SqlConnection(connectionString))
                {
                    connection.Open();
                    using (var command = new SqlCommand(insertSql, connection))
                    {
                        foreach (DataRow row in dataTable.Rows)
                        {
                            foreach (DataColumn col in dataTable.Columns)
                            {
                                if (col.DataType == typeof(string))
                                {
                                    int maxLength = columnMaxLengths[dataTable.Columns.IndexOf(col)];
                                    string value = row[col].ToString();
                                    if (value.Length > maxLength)
                                    {
                                        value = value.Substring(0, maxLength);
                                    }
                                    command.Parameters.Add($"@{col.ColumnName}", SqlDbType.NVarChar, -1).Value = row[col];
                                }
                                else
                                {
                                    command.Parameters.AddWithValue($"@{col.ColumnName}", row[col]);
                                }
                            }
                            System.Diagnostics.Debug.WriteLine(command.CommandText);
                            command.ExecuteNonQuery();
                            command.Parameters.Clear();
                        }
                    }
                }


            }

        }

        
        // Open a connection to the SQL Server database
        using (SqlConnection connection = new SqlConnection(connectionString))
        {
            connection.Open();

            // Create a SQL command to insert data into the table
            SqlCommand command = new SqlCommand();

            // Set the connection for the SQL command
            command.Connection = connection;

            // Add the table name to the SQL command
            command.CommandText = $"INSERT INTO {tableName} ({string.Join(",", columnNames)}) VALUES ";

            command.CommandText += "(";
            // Add a parameter for each column in the table
            for (int i = 0; i < columnNames.Length; i++)
            {
                command.CommandText += $"{dataTable.Rows}";
                command.Parameters.Add($"{columnNames[i]}", SqlDbType.VarChar);
                if (i < columnNames.Length - 1)
                {
                    command.CommandText += ", ";
                }
            }

            command.CommandText += ")";
            // Print the SQL command before executing it
            System.Diagnostics.Debug.WriteLine(command.CommandText);

        

    }
}
        
    
    protected void btnUpload_Click(object sender, EventArgs e)
    {
        using (XLWorkbook workBook = new XLWorkbook(fileUpload.PostedFile.InputStream))
        {
            IXLWorksheet workSheet = workBook.Worksheet(1);
            DataTable dt = new DataTable();
            var table = "Progetti";
            bool firstRow = true;
            foreach (IXLRow row in workSheet.Rows())
            {
                if (firstRow)
                {
                    foreach (IXLCell cell in row.Cells())
                    {
                        dt.Columns.Add(cell.Value.ToString());
                    }
                    firstRow = false;
                }
                else
                {
                    dt.Rows.Add();
                    int i = 0;
                    foreach (IXLCell cell in row.Cells())
                    {
                        dt.Rows[dt.Rows.Count - 1][i] = cell.Value.ToString();
                        i++;
                    }
                }

                GridView1.DataSource = dt;
                GridView1.DataBind();
            }
            string cs = ConfigurationManager.ConnectionStrings["connectionString"].ConnectionString;
            using (var bulkCopy = new SqlBulkCopy(cs))
            {

                foreach (DataColumn col in dt.Columns)
                {
                    bulkCopy.ColumnMappings.Add(col.ColumnName, col.ColumnName);
                }
                bulkCopy.DestinationTableName = table;
                bulkCopy.WriteToServer(dt);
            }
        }
    */

    /*

        private List<Progetto> GetProjects()
        {
            var progetti = new List<Progetto>();

            var connectionString = "Data Source=VT-LEONFAB;Initial Catalog=ProgettiSocialiCustom;User ID=sa;Password=6LGmG!iDrzq4";
            var query = "SELECT * FROM Progetti ORDER BY nProtocollo";
            var dataTable = new DataTable();



            using (var connection = new SqlConnection(connectionString))
            {
                connection.Open();
                SqlCommand command = new SqlCommand(query, connection);
                using (SqlDataReader reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        Progetto progetto = new Progetto
                        {
                            Anno = reader.GetString(0),
                            nProtocollo = reader.GetInt32(1),
                            dataInserimento = reader.GetDateTime(2),
                            Tipologia = reader.GetString(3),
                            Stato = reader.GetString(4),
                            Ambito = reader.GetString(5),
                            Soggetti = reader.GetString(6),
                            Titolo = reader.GetString(7),
                            dataInizio = reader.GetDateTime(8),
                            dataFine = reader.GetDateTime(9)
                        };
                        progetti.Add(progetto);
                    }
                }
            }

            return progetti;
        }

        private string GetProjectsJson()
        {
            var projects = GetProjects();
            var json = JsonConvert.SerializeObject(projects);
            System.Diagnostics.Debug.WriteLine(json);

            return json;
        }

        private string GetProjectsJson()
        {
            Dictionary<string, Progetto> progetti = new Dictionary<string, Progetto>();

            


            var json = JsonConvert.SerializeObject(progetti);

            return json;
        }
        */


    /*
    for (int i = 0; i < table.Rows.Count; i++)
    {
        //System.Diagnostics.Debug.WriteLine("ciao"+table.Rows[i][1]);
        for (int j = 0; j < table.Rows[i].ItemArray.Length; j++)
        {
            //System.Diagnostics.Debug.WriteLine(table.Rows[i].ItemArray.GetValue(j));
            System.Diagnostics.Debug.WriteLine("ciao" + table.Rows[i][j]);
        }
    }

    //aggiunge righe alla tabella in base ai dati del foglio di lavoro
    for (int i = 2; i <= worksheet.LastRowUsed().RowNumber(); i++)
    {
        var row = worksheet.Row(i);
        System.Diagnostics.Debug.WriteLine(row);
        var data = row.Cells().Select(cell => cell.Value.ToString()).ToArray();


        //controlla che ogni cella sia piena e che il valore sia del tipo corretto
        for (var j = 0; j < table.Columns.Count; j++)
        {
            var columnType = table.Columns[j].DataType;
            var cellValue = data[j];

            if (cellValue == null || cellValue == "")
            {
                //se la cella è vuota, aggiungi un errore alla riga
                data[j] = "Errore: il campo è vuoto";
            }
            else if (columnType == typeof(int) && !int.TryParse(cellValue, out int _))
            {
                //se il tipo di colonna è int ma il valore non può essere convertito in int, aggiungi un errore alla riga
                data[j] = "Errore: il campo non è un numero intero";
            }
            else if (columnType == typeof(DateTime) && !DateTime.TryParse(cellValue, out DateTime _))
            {
                //se il tipo di colonna è DateTime ma il valore non può essere convertito in DateTime, aggiungi un errore alla riga
                data[j] = "Errore: il campo non è una data";
            }
        }


        table.Rows.Add(data);
    }
    */

    /* inserimento db
    var connectionString = "Data Source=VT-LEONFAB;Initial Catalog=ProgettiSocialiCustom;User ID=sa;Password=6LGmG!iDrzq4";
    using (var connection = new SqlConnection(connectionString))
    {
        var command = new SqlCommand("INSERT INTO Progetti VALUES (@Anno, @nProtocollo, @dataInserimento, @Tipologia, @Stato, @Ambito, @Soggetti, @Titolo, @dataInizio, @dataFine, @Errori)", connection);
        command.Parameters.Add("@Anno", SqlDbType.NChar);
        command.Parameters.Add("@nProtocollo", SqlDbType.Int);
        command.Parameters.Add("@dataInserimento", SqlDbType.Date);
        command.Parameters.Add("@Tipologia", SqlDbType.NChar);
        command.Parameters.Add("@Stato", SqlDbType.NChar);
        command.Parameters.Add("@Ambito", SqlDbType.VarChar);
        command.Parameters.Add("@Soggetti", SqlDbType.VarChar);
        command.Parameters.Add("@Titolo", SqlDbType.VarChar);
        command.Parameters.Add("@dataInizio", SqlDbType.Date);
        command.Parameters.Add("@dataFine", SqlDbType.Date);
        command.Parameters.Add("@Errori", SqlDbType.VarChar);

        connection.Open();
        int counter = 0;
        foreach (DataRow row in table.Rows)
        {
            command.Parameters["@Anno"].Value = row["ANNO"].ToString();
            command.Parameters["@nProtocollo"].Value = row["NR. PROTOCOLLO"].ToString();
            command.Parameters["@dataInserimento"].Value = row["DATA INSERIMENTO"].ToString();
            command.Parameters["@Tipologia"].Value = row["TIPOLOGIA"].ToString();
            command.Parameters["@Stato"].Value = row["STATO"].ToString();
            command.Parameters["@Ambito"].Value = row["AMBITO D'INTERVENTO"].ToString();
            command.Parameters["@Soggetti"].Value = row["SOGGETTI DESTINATARI\n"].ToString();
            command.Parameters["@Titolo"].Value = row["TITOLO INIZIATIVA "].ToString();
            command.Parameters["@dataInizio"].Value = row["DATA INIZIO\n(GG/MM/AA)"].ToString();
            command.Parameters["@dataFine"].Value = row["DATA FINE\n(GG/MM/AA)"].ToString();
            command.Parameters["@Errori"].Value = row.ToString();

            counter++;
            command.ExecuteNonQuery();
        }

        //System.Diagnostics.Debug.WriteLine("Numero di righe interessate: " + counter);
        if (counter == 1)
        {
            lblRowCount.Text = "É stata interessata " + counter + " riga.";
        }
        else
        {
            lblRowCount.Text = "Sono state interessate " + counter + " righe.";
        }


    }//qua trasferisce i dati al database, ogni volta bisogna svuotare il database per provare
    */

    /*
            // recupera i dati e li visualizza nella tabella
            var projectsJson = GetProjectsJson();
            var projects = JArray.Parse(projectsJson);

            foreach (var project in projects)
            {
                var row = new HtmlTableRow();
                row.Cells.Add(new HtmlTableCell() { InnerHtml = project["Anno"].ToString() });
                row.Cells.Add(new HtmlTableCell() { InnerHtml = project["nProtocollo"].ToString() });
                row.Cells.Add(new HtmlTableCell() { InnerHtml = project["dataInserimento"].ToString() });
                row.Cells.Add(new HtmlTableCell() { InnerHtml = project["Tipologia"].ToString() });
                row.Cells.Add(new HtmlTableCell() { InnerHtml = project["Stato"].ToString() });
                row.Cells.Add(new HtmlTableCell() { InnerHtml = project["Ambito"].ToString() });
                row.Cells.Add(new HtmlTableCell() { InnerHtml = project["Soggetti"].ToString() });
                row.Cells.Add(new HtmlTableCell() { InnerHtml = project["Titolo"].ToString() });
                row.Cells.Add(new HtmlTableCell() { InnerHtml = project["dataInizio"].ToString() });
                row.Cells.Add(new HtmlTableCell() { InnerHtml = project["dataFine"].ToString() });
                row.Cells.Add(new HtmlTableCell() { InnerHtml = project["Errori"].ToString() });

                projectsTable.Rows.Add(row);
            }
            // nasconde lo spinner
            ClientScript.RegisterStartupScript(GetType(), "hideSpinner", "<script>hideSpinner();</script>");
            
            projectsTable.Visible = true;
            */

    /*
     <!--
    <asp:Label ID="validationErrorsLabel" runat="server" />
    <asp:Label ID="lblRowCount" runat="server" />
    <asp:Panel runat="server" ID="panelData" CssClass="panel panel-default" Visible="false">
         i dati vengono mostrati qui -->
    <!--
    <div class="table-responsive">
        <table class="table table-hover table-bordered" id="projectsTable" runat="server">
            <thead>
                <tr>
                    <th>Anno</th>
                    <th>Nr. Protocollo</th>
                    <th>Data Inserimento</th>
                    <th>Tipologia</th>
                    <th>Stato</th>
                    <th>Ambito d'intervento</th>
                    <th>Soggetti Destinatari</th>
                    <th>Titolo Iniziativa</th>
                    <th>Data Inizio&#10;(GG/MM/AA)</th>
                    <th>Data Fine&#10;(GG/MM/AA)</th>
                    <th>Errori</th>
                </tr>
            </thead>
        </table>
    </div>
    <asp:LinkButton ID="LinkButton2" runat="server" OnClick="btnReset_Click" CssClass="btn btn-secondary"><i class="bi bi-trash"></i>&nbsp;Reset</asp:LinkButton>

    </asp:Panel>
    -->
     */

    /*
       private string GetProjectsJson()
       {
           var connectionString = "Data Source=VT-LEONFAB;Initial Catalog=ProgettiSocialiCustom;User ID=sa;Password=6LGmG!iDrzq4";
           var query = $"SELECT * FROM Progetti order by nProtocollo";
           var dataTable = new DataTable();

           using (var connection = new SqlConnection(connectionString))
           {
               var adapter = new SqlDataAdapter(query, connection);
               adapter.Fill(dataTable);
           }

           var jsonSettings = new JsonSerializerSettings();
           jsonSettings.DateFormatString = "dd/MM/yyyy";

           var json = JsonConvert.SerializeObject(dataTable, jsonSettings);
           System.Diagnostics.Debug.WriteLine(json);

           return json;
       }*/

    /*
        <!--
        <script type = "text/javascript" >
        function Get()
    {
            $.GetProjectsJson("api/projects",
                function(data) {
                    $('#products').empty(); // Clear the table body.

                    // Loop through the list of products.
                    $.each(data, function(key, val) {
                // Add a table row for the product.
                var row = '<td>' + val.Anno + '</td><td>' + val.nProtocollo + +'</td><td>' + val.dataInserimento + '</td><td>' + val.Tipologia +
                    "</td><td>" + item.Stato + "</td><td>" + item.Ambito + "</td><td>" +
                    val.Soggetti + "</td><td>" + val.Titolo + "</td><td>" + val.dataInizio +
                    "</td><td>" + val.dataFine + "</td><td>" + val.Errori + '</td>';
                        $('<tr/>', { html: row })  // Append the name.
                            .appendTo($('#projects'));
            });
        });
    }

        $(document).ready(GetProjectsJson);
</script> 

    <script type = "text/javascript" >
        // make an AJAX call to retrieve the data from the server in JSON format
        $.ajax({
    url: 'MyEndpoint.aspx',
            method: 'GET',
            dataType: 'json',
            success: function(data) {
                // loop through the data and append it to the table
                $.each(data, function(i, item) {
                var row = "<tbody><tr><td>" + item.Anno + "</td><td>" + item.nProtocollo +
                    "</td><td>" + item.dataInserimento + "</td><td>" + item.Tipologia +
                    "</td><td>" + item.Stato + "</td><td>" + item.Ambito + "</td><td>" +
                    item.Soggetti + "</td><td>" + item.Titolo + "</td><td>" + item.dataInizio +
                    "</td><td>" + item.dataFine + "</td><td>" + item.Errori + "</td></tr></tbody>";
                    $('#projectsTable').append(row);
            });
        }
    });
    </script>

    <script type = "text/javascript" >
        $.ajax({
    url: 'MyEndpoint.aspx',
            method: 'GET',
            dataType: 'json',
            success: function(data) {
                // loop through the data and append it to the table
                $.each(data, function(i, item) {
                var row = "<tbody><tr";
                if (!item.Titolo)
                {
                    row += " class='bg-warning'";
                }
                row += "><td>" + item.Anno + "</td><td>" + item.nProtocollo +
                    "</td><td>" + item.dataInserimento + "</td><td>" + item.Tipologia +
                    "</td><td>" + item.Stato + "</td><td>" + item.Ambito + "</td><td>" +
                    item.Soggetti + "</td><td>" + item.Titolo + "</td><td>" + item.dataInizio +
                    "</td><td>" + item.dataFine + "</td></tr></tbody>";
                    $('#projectsTable').append(row);
            });
        }
    });

    </script>
        */

    /*
     foreach (var row in range.RowsUsed().Skip(1))
            {
                var data = new List<string>();
                bool hasEmptyCell = false;
                string columnName = string.Empty;
                for (int i = 1; i <= table.Columns.Count - 1; i++)
                {
                    var cell = row.Cell(i);
                    if (cell.IsEmpty())
                    {
                        data.Add("");
                        columnName = table.Columns[i - 1].ColumnName;
                        hasEmptyCell = true;
                    }
                    else
                    {
                        data.Add(cell.Value.ToString());
                    }

                }

                if (hasEmptyCell)
                {
                    table.Rows.Add(data.Concat(new string[] { "Il campo " + columnName + " è vuoto" }).ToArray());
                }
                else
                {
                    table.Rows.Add(data.ToArray());
                }

            }
     */


    /*
      protected void btnUpload_Click(object sender, EventArgs e)
        {


            if (fileUpload.HasFile)
            {
                // mostra i dati
                panelData.Visible = true;
                // ottiene il DataTable dal file Excel
                var table = fromExcelToDataTable();
                // valida il DataTable appena creato
                ValidateDataTable(table);

                // inserisce i dati nel database se non ci sono errori
                if (table.AsEnumerable().All(row => string.IsNullOrEmpty(row["Errori"].ToString())))
                {
                    fromDataTabletoDatabase(table);

                    // nasconde lo spinner
                    ClientScript.RegisterStartupScript(GetType(), "hideSpinner", "<script>hideSpinner();</script>");

                    // visualizza il GridView
                    JsonToGridView();
                    projectsGridView.Visible = true;
                    successLabel.Text = "Importazione completata con successo.";
                    successLabel.Visible = true;
                    errorLabel.Visible = false;
                }
                else
                {
                    // Mostra gli errori all'utente
                    var errorRows = table.AsEnumerable().Where(row => !string.IsNullOrEmpty(row["Errori"].ToString()));
                    var errorMessages = errorRows.Select(row => row["Errori"].ToString()).Distinct();
                    errorLabel.Text = string.Join("<br>", errorMessages);
                    successLabel.Visible = false;
                    errorLabel.Visible = true;
                }
            }
            else
            {
                // gestione dell'errore se il file non è stato caricato
                errorLabel.Text = "Selezionare un file Excel da caricare.";
                successLabel.Visible = false;
                errorLabel.Visible = true;
            }
        }
     */


    /*
     foreach (var row in range.RowsUsed().Skip(1))
                {
                    var data = new List<object>();
                    bool hasEmptyCell = false;
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
                        else
                        {
                            if (expectedType == typeof(DateTime) && cell.DataType == XLDataType.DateTime)
                            {
                                value = cell.GetDateTime();
                            }
                            else
                            {
                                // tenta di convertire il valore della cella nel tipo di dato atteso
                                value = Convert.ChangeType(cell.Value, expectedType);
                            }
                        }

                        data.Add(value);
                    }

                    if (hasEmptyCell)
                    {
                        // se la riga contiene almeno una cella vuota, segnala un errore
                        table.Rows.Add(data.Concat(new string[] { "Il campo " + columnName + " è vuoto" }).ToArray());
                    }
                    else
                    {
                        // altrimenti aggiunge la riga al DataTable
                        table.Rows.Add(data.ToArray());
                    }
                }
     */

    /*
     foreach (var row in range.RowsUsed().Skip(1))
                {
                    var data = new List<object>();
                    bool hasEmptyCell = false;
                    bool hasWrongType = false;
                    string columnName = string.Empty;

                    for (int i = 1; i <= table.Columns.Count - 1; i++)
                    {
                        var cell = row.Cell(i + 1);
                        object value = null;
                        var expectedType = table.Columns[i].DataType;
                        var cellValue = cell.Value;
                        var columnType = table.Columns[i - 1].DataType;
                        if (cell.IsEmpty())
                        {
                            value = DBNull.Value;
                            columnName = table.Columns[i - 1].ColumnName;
                            hasEmptyCell = true;
                        }
                        else
                        {
                            if (expectedType != columnType)
                            {
                                value = Convert.ChangeType(cellValue, columnType);
                            }
                            else
                            {
                                value = DBNull.Value;
                                columnName = table.Columns[i - 1].ColumnName;
                                hasWrongType = true;
                                //table.Rows.Add(data.Concat(new object[] { "Il campo " + columnName + " contiene un valore non valido" }).ToArray());
                                //break;
                            }

                            if (expectedType == typeof(DateTime) && cell.DataType == XLDataType.DateTime)
                            {
                                value = cell.GetDateTime();
                            }
                            else
                            {
                                // tenta di convertire il valore della cella nel tipo di dato atteso
                                value = Convert.ChangeType(cellValue, expectedType);
                            }
                        }
                        data.Add(value);
                    }

                    if (hasEmptyCell)
                    {
                        table.Rows.Add(data.Concat(new string[] { "Il campo " + columnName + " è vuoto" }).ToArray());
                    }
                    else if (hasWrongType)
                    {
                        table.Rows.Add(data.Concat(new object[] { "Il campo " + columnName + " contiene un valore non valido" }).ToArray());
                    }
                    else
                    {
                        table.Rows.Add(data.ToArray());
                    }


                }
     */
}