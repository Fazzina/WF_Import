using ClosedXML.Excel;
using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace WF_Import
{
    public partial class About : Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            ShowButton.Visible = false;

            if (checkEmptyTable() == 0)
            {
                panelData.Visible = false;
            }
            else
            {
                panelData.Visible = true;
            }

            if (SearchContron.Text.Length > 0)
            {
                SqlDataSource1.SelectCommand = "SELECT * FROM [Progetti] WHERE (nProtocollo LIKE N'%" + SearchContron.Text + "%' or Anno LIKE N'%" + SearchContron.Text + "%' or dataInserimento LIKE N'%" + SearchContron.Text + "%' or Tipologia LIKE N'%" + SearchContron.Text + "%' or Stato LIKE N'%" + SearchContron.Text + "%' or Ambito LIKE N'%" + SearchContron.Text + "%' or Soggetti LIKE N'%" + SearchContron.Text + "%' or Titolo LIKE N'%" + SearchContron.Text + "%' or dataInizio LIKE N'%" + SearchContron.Text + "%' or dataFine LIKE N'%" + SearchContron.Text + "%' ) ORDER BY nProtocollo";
            }
            else
            {
                //SqlDataSource1.SelectCommand = "SELECT * ORDER BY [Nr. Protocollo]";
            }



        }
        public string connectionString()
        {
            return ConfigurationManager.ConnectionStrings["connectionString"].ConnectionString;
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
                    return result; // se il risultato è uguale a zero, la tabella è vuota
                }
                catch (Exception)
                {
                    throw;
                }
            }
        }


        protected void OnRowDataBound(object sender, GridViewRowEventArgs e)
        {
            // Controlla nel caso ci sia un campo data, formatta l'output nel frontend
            if (e.Row.RowType == DataControlRowType.DataRow)
            {
                foreach (TableCell cell in e.Row.Cells)
                {
                    DateTime dateValue;
                    if (DateTime.TryParse(cell.Text, out dateValue))
                    {
                        cell.Text = dateValue.ToString("dd/MM/yyyy");
                    }
                }
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

            if (checkEmptyTable() == 0)
            {
                ShowButton.Visible = false;
            }
            else
            {
                ShowButton.Visible = true;
            }

            panelData.Visible = false;
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
        protected void OnSubmitButtonClick(object sender, EventArgs e)
        {
            projectsGridView.DataBind();
        }

        protected void ChkEmpty_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox chkstatus = (CheckBox)sender;
            GridViewRow row = (GridViewRow)chkstatus.NamingContainer;
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

        protected void ChkHeader_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox chkheader = (CheckBox)projectsGridView.HeaderRow.FindControl("ChkHeader");
            foreach (GridViewRow row in projectsGridView.Rows)
            {
                CheckBox chkrow = (CheckBox)row.FindControl("ChkEmpty");
                if (chkheader.Checked == true)
                {
                    chkrow.Checked = true;
                }
                else
                {
                    chkrow.Checked = false;
                }
            }
        }

        protected void projectsGridView_RowEditing(object sender, GridViewEditEventArgs e)
        {
            projectsGridView.EditIndex = e.NewEditIndex;
            projectsGridView.DataBind();
        }


        protected void projectsGridView_RowUpdating(object sender, GridViewUpdateEventArgs e)
        {
            try
            {
                int nPro = Convert.ToInt32(projectsGridView.DataKeys[e.RowIndex].Value.ToString());

                string anno1 = ((TextBox)projectsGridView.Rows[e.RowIndex].Cells[1].Controls[0]).Text;
                string prot = ((TextBox)projectsGridView.Rows[e.RowIndex].Cells[2].Controls[0]).Text;
                DateTime dataIns = DateTime.ParseExact(((TextBox)projectsGridView.Rows[e.RowIndex].Cells[3].Controls[0]).Text, "dd/MM/yyyy HH:mm:ss", CultureInfo.InvariantCulture);
                string tipologia1 = ((TextBox)projectsGridView.Rows[e.RowIndex].Cells[4].Controls[0]).Text;
                string stato1 = ((TextBox)projectsGridView.Rows[e.RowIndex].Cells[5].Controls[0]).Text;
                string ambito1 = ((TextBox)projectsGridView.Rows[e.RowIndex].Cells[6].Controls[0]).Text;
                string soggetti1 = ((TextBox)projectsGridView.Rows[e.RowIndex].Cells[7].Controls[0]).Text;
                string titolo1 = ((TextBox)projectsGridView.Rows[e.RowIndex].Cells[8].Controls[0]).Text;
                DateTime dataIn = DateTime.ParseExact(((TextBox)projectsGridView.Rows[e.RowIndex].Cells[9].Controls[0]).Text, "dd/MM/yyyy HH:mm:ss", CultureInfo.InvariantCulture);
                DateTime dataFn = DateTime.ParseExact(((TextBox)projectsGridView.Rows[e.RowIndex].Cells[10].Controls[0]).Text, "dd/MM/yyyy HH:mm:ss", CultureInfo.InvariantCulture);

                using (SqlConnection con = new SqlConnection(connectionString()))
                {
                    con.Open();
                    SqlCommand cmd = new SqlCommand("update Progetti set Anno='" + anno1 + "', nProtocollo='" + prot + "', dataInserimento= CONVERT(datetime,'" + dataIns + "', 103), Tipologia='" + tipologia1 + "', Stato='" + stato1 + "', Ambito='" + ambito1 + "', Soggetti='" + soggetti1 + "', Titolo='" + titolo1 + "',dataInizio= CONVERT(datetime,'" + dataIn + "', 103), dataFine= CONVERT(datetime,'" + dataFn + "', 103) where nProtocollo='" + nPro + "'", con);
                    SqlDataSource1.UpdateCommand = "update Progetti set Anno='" + anno1 + "', nProtocollo='" + prot + "', dataInserimento= CONVERT(datetime,'" + dataIns + "', 103), Tipologia='" + tipologia1 + "', Stato='" + stato1 + "', Ambito='" + ambito1 + "', Soggetti='" + soggetti1 + "', Titolo='" + titolo1 + "',dataInizio= CONVERT(datetime,'" + dataIn + "', 103), dataFine= CONVERT(datetime,'" + dataFn + "', 103) where nProtocollo='" + nPro + "'";
                    //SqlCommand cmd = new SqlCommand("update Progetti set Anno='" + anno1 + "', nProtocollo='" + prot + "', Tipologia='" + tipologia1 + "', Stato='" + stato1 + "', Ambito='" + ambito1 + "', Soggetti='" + soggetti1 + "', Titolo='" + titolo1 + "' where nProtocollo='" + nPro + "'", con);
                    //SqlDataSource1.UpdateCommand = "update Progetti set Anno='" + anno1 + "', nProtocollo='" + prot + "', Tipologia='" + tipologia1 + "', Stato='" + stato1 + "', Ambito='" + ambito1 + "', Soggetti='" + soggetti1 + "', Titolo='" + titolo1 + "' where nProtocollo='" + nPro + "'";
                    int t = cmd.ExecuteNonQuery();
                    if (t > 0)
                    {
                        //Response.Write("<script>alert('Data ha updated')</script>");
                        projectsGridView.EditIndex = -1;
                        SqlDataSource1.DataBind();
                        projectsGridView.DataBind();
                    }
                }
            }
            catch (Exception)
            {
            }
        }



        protected void projectsGridView_RowCancelingEdit(object sender, GridViewCancelEditEventArgs e)
        {
            projectsGridView.EditIndex = -1;
            projectsGridView.DataBind();
        }

        protected void projectsGridView_RowDeleting(object sender, GridViewDeleteEventArgs e)
        {
            int nPro = Convert.ToInt32(projectsGridView.DataKeys[e.RowIndex].Value.ToString());

            using (SqlConnection con = new SqlConnection(connectionString()))
            {
                con.Open();
                SqlCommand cmd = new SqlCommand("delete from Progetti where nProtocollo='" + nPro + "'", con);
                SqlDataSource1.DeleteCommand = "delete from Progetti where nProtocollo='" + nPro + "'";
                int t = cmd.ExecuteNonQuery();
                if (t > 0)
                {
                    //Response.Write("<script>alert('Data ha updated')</script>");
                    projectsGridView.EditIndex = -1;
                    projectsGridView.DataBind();
                }
            }

            if (checkEmptyTable() == 0)
            {
                panelData.Visible = false;
            }
        }

        protected void btnDelete_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < projectsGridView.Rows.Count; i++)
            {
                CheckBox chkdelete = (CheckBox)projectsGridView.Rows[i].Cells[0].FindControl("ChkEmpty");
                if (chkdelete.Checked)
                {
                    int nPro = Convert.ToInt32(projectsGridView.Rows[i].Cells[2].Text);
                    using (SqlConnection con = new SqlConnection(connectionString()))
                    {
                        con.Open();
                        SqlCommand cmd = new SqlCommand("delete from Progetti where nProtocollo='" + nPro + "'", con);
                        SqlDataSource1.DeleteCommand = "delete from Progetti where nProtocollo='" + nPro + "'";
                        cmd.ExecuteNonQuery();
                        con.Close();
                    }
                }

            }

            if (checkEmptyTable() == 0)
            {
                panelData.Visible = false;
            }

            projectsGridView.DataBind();
        }

        protected void btnInsert_Click(object sender, EventArgs e)
        {
            // mostra il popup
            ScriptManager.RegisterStartupScript(this, GetType(), "Popup", "showPopup();", true);
        }

        protected void btnSubmit_Click(object sender, EventArgs e)
        {
            using (SqlConnection con = new SqlConnection(connectionString()))
            {
                con.Open();

                SqlCommand cmd = new SqlCommand("INSERT INTO Progetti VALUES ('" + anno.Value + "','" + protocollo.Value + "','" + datainserimento.Value + "','" + tipologia.Value + "','" + stato.Value + "','" + ambito.Value + "','" + soggetti.Value + "','" + titolo.Value + "','" + datainizio.Value + "','" + datafine.Value + "')", con);
                SqlDataSource1.DeleteCommand = "INSERT INTO Progetti VALUES ('" + anno.Value + "','" + protocollo.Value + "','" + datainserimento.Value + "','" + tipologia.Value + "','" + stato.Value + "','" + ambito.Value + "','" + soggetti.Value + "','" + titolo.Value + "','" + datainizio.Value + "','" + datafine.Value + "')";

                int t = cmd.ExecuteNonQuery();
                if (t > 0)
                {
                    projectsGridView.DataBind();
                }

            }

            if (checkEmptyTable() == 0)
            {
                panelData.Visible = false;
            }
            else
            {
                panelData.Visible = true;
            }
        }

        protected void btnExport_Click(object sender, EventArgs e)
        {
            DataTable dataTable = new DataTable();

            dataTable.Columns.Add("Anno", typeof(string));
            dataTable.Columns.Add("Nr. Protocollo", typeof(int));
            dataTable.Columns.Add("Data Inserimento", typeof(DateTime));
            dataTable.Columns.Add("Tipologia", typeof(string));
            dataTable.Columns.Add("Stato", typeof(string));
            dataTable.Columns.Add("Ambito d'intervento", typeof(string));
            dataTable.Columns.Add("Soggetti Destinatari", typeof(string));
            dataTable.Columns.Add("Titolo Iniziativa", typeof(string));
            dataTable.Columns.Add("Data Inizio (GG/MM/AA)", typeof(DateTime));
            dataTable.Columns.Add("Data Fine (GG/MM/AA)", typeof(DateTime));


            foreach (GridViewRow row in projectsGridView.Rows)
            {
                CheckBox chkRow = (row.Cells[0].FindControl("ChkEmpty") as CheckBox);
                if (chkRow.Checked)
                {
                    DataRow dr = dataTable.NewRow();
                    for (int i = 1; i < row.Cells.Count - 1; i++)
                    {
                        dr[i - 1] = row.Cells[i].Text;
                    }
                    dataTable.Rows.Add(dr);
                }
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
            // Esporta datatable in un file excel
            var fileName = "Progetti.xlsx";
            var filePath = Server.MapPath("~/") + fileName;

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("MAPPATURA");
                worksheet.Cell(1, 1).InsertTable(dataTable);
                workbook.SaveAs(filePath);
            }

            // Invia il file come allegato nella risposta 
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.AppendHeader("Content-Disposition", $"attachment; filename={fileName}");
            Response.TransmitFile(filePath);
            Response.End();
        }
    }
}