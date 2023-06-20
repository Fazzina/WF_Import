<%@ Page Title="View" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="View.aspx.cs" Inherits="WF_Import.About" %>

<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">

    <style>
        #MainContent_txtSearch, #MainContent_btnSearch {
            color: black;
        }

        #MainContent_SearchContron {
            color: black;
            margin-left: 3px;
            margin-bottom: 3px;
            margin-top: 3px;
            height: 30px;
            border-radius: 5px;
        }

        #MainContent_SearchButton {
            height: 23px;
            margin-left: -18px;
        }

        #MainContent_ResetButton,
        #MainContent_HideButton {
            float: right;
        }

        #MainContent_btnDelete {
            margin-left: 3px;
            margin-bottom: 3px;
            margin-top: 3px;
        }

        #MainContent_btnInsert {
            margin-bottom: 35px;
            background: linear-gradient(90deg, #D7E1EC 0%, #FFFFFF 100%);
            color: black;
        }

            #MainContent_btnInsert:hover {
                background: linear-gradient(90deg, #D7E1EC 0%, #E2E6EA 100%);
            }

        .form-group {
            margin-bottom: 20px;
            color: black;
        }

        label {
            display: block;
            font-weight: bold;
            margin-bottom: 5px;
        }

        input[type="text"], input[type="number"], input[type="date"] {
            padding: 10px;
            border-radius: 5px;
            border: 1px solid #ccc;
            font-size: 16px;
            width: 220px;
        }

        .left {
            float: left;
            margin-right: 20px;
            margin-top: 20px;
        }

        .right {
            float: right;
            margin-left: 20px;
        }
    </style>

    <script>
const { queryselector } = require("modernizr");

        function showPopup() {
            document.getElementById('popup').style.display = 'block';
        }

        function closePopup() {
            document.getElementById('popup').style.display = 'none';
        }

        function setDate() {
            var datainizio = document.getElementById("MainContent_datainizio");
            var datafine = document.getElementById("MainContent_datafine");

            datafine.setAttribute("min", datainizio.value)
            // imposta la data minima della data fine uguale alla data di inizio    
        };        

        function checkEmpty() {
            if (document.querySelector('.check input').value == "") {
                alert("please enter some thing")
                document.querySelector('.check input').focus();
                event.preventDefault();
                return false;
            }

            alert("validations passed");
            return true;
        }

    </script>


    <h2><%: Title %></h2>
    <p>If you have already uploaded an Excel file, the data will be displayed here.</p>
    <br>

    <asp:LinkButton ID="btnInsert" CssClass="btn btn-light" OnClick="btnInsert_Click" runat="server"><img src="databaseadd.png"/>&nbsp;&nbsp;Insert data in DB</asp:LinkButton>
    <div id="popup" class="popup">
        <div style="background: linear-gradient(90deg, #D7E1EC 0%, #FFFFFF 100%);" class="popup-content">
            <span class="popup-close" onclick="closePopup()">&times;</span>
            <div class="popup-image">
                <img runat="server" id="popupImg" src="db.png" />
            </div>
            <div class="check" style="overflow: auto;">
                <div class="left">
                    <div class="form-group" style="width: 50%; margin-top: 20px;">
                        <label for="anno">Anno:</label>
                        <input type="number" id="anno" name="anno" min="1900" max="2099" maxlength="4" runat="server" required>
                    </div>
                    <div class="form-group" style="width: 50%;">
                        <label for="protocollo">Nr. Protocollo:</label>
                        <input type="number" id="protocollo" runat="server" name="protocollo" required>
                    </div>
                    <div class="form-group" style="width: 50%;">
                        <label for="datainserimento">Data Inserimento:</label>
                        <input type="date" id="datainserimento" runat="server" name="datainserimento" required>
                    </div>
                    <div class="form-group" style="width: 50%;">
                        <label for="tipologia">Tipologia:</label>
                        <input type="text" name="tipologia" id="tipologia" runat="server" required />
                    </div>
                    <div class="form-group" style="width: 50%;">
                        <label for="stato">Stato:</label>
                        <input type="text" name="stato" id="stato" runat="server" required />
                    </div>
                </div>

                <div class="right">
                    <div class="form-group" style="width: 50%; margin-top: 20px;">
                        <label for="ambito">Ambito d'intervento:</label>
                        <input type="text" name="ambito" id="ambito" runat="server" required />
                    </div>
                    <div class="form-group" style="width: 50%;">
                        <label for="soggetti">Soggetti Destinatari:</label>
                        <input type="text" name="soggetti" id="soggetti" runat="server" required />

                    </div>
                    <div class="form-group" style="width: 50%;">
                        <label for="titolo">Titolo Iniziativa:</label>
                        <input type="text" name="titolo" id="titolo" runat="server" required />

                    </div>
                    <div class="form-group" style="width: 50%;">
                        <label for="datainizio">Data Inizio:</label>
                        <input type="date" id="datainizio" runat="server" onchange="setDate();" name="datainizio" required>
                    </div>
                    <div class="form-group" style="width: 50%;">
                        <label for="datafine">Data Fine:</label>
                        <input type="date" min="" id="datafine" runat="server" name="datafine" required>
                    </div>
                </div>
            </div>

            <asp:LinkButton ID="btnSubmit" CssClass="btn btn-success" Font-Bold="true" OnClientClick="checkEmpty();" OnClick="btnSubmit_Click" runat="server"><img style="width: 14px;" src="thick.png"/>&nbsp;INSERT</asp:LinkButton>
        </div>
    </div>

    <asp:Panel runat="server" ID="panelData" CssClass="panel panel-default" Visible="false">
        <asp:TextBox runat="server" placeholder="Search..." ID="SearchContron" />
        <asp:ImageButton ImageUrl="~/search.png" ID="SearchButton" OnClick="OnSubmitButtonClick" runat="server" />
        <div class="table-responsive">
            <asp:SqlDataSource ID="SqlDataSource1" DataSourceMode="DataSet" runat="server" ConnectionString="<%$ ConnectionStrings:connectionString%>"
                SelectCommand="SELECT * FROM Progetti ORDER BY nProtocollo"></asp:SqlDataSource>
            <asp:GridView ID="projectsGridView" SelectionMode="Multiple" DataSourceID="SqlDataSource1" OnRowDataBound="OnRowDataBound" AutoGenerateColumns="False" CssClass="table table-hover table-bordered" runat="server" DataKeyNames="nProtocollo" AllowEditing="True" OnRowEditing="projectsGridView_RowEditing" OnRowUpdating="projectsGridView_RowUpdating" OnRowCancelingEdit="projectsGridView_RowCancelingEdit" OnRowDeleting="projectsGridView_RowDeleting">
                <Columns>
                    <asp:TemplateField>
                        <HeaderTemplate>
                            <asp:CheckBox ID="ChkHeader" runat="server" Text="Select All" AutoPostBack="true" OnCheckedChanged="ChkHeader_CheckedChanged" />
                        </HeaderTemplate>
                        <ItemTemplate>
                            <asp:CheckBox ID="ChkEmpty" runat="server" OnCheckedChanged="ChkEmpty_CheckedChanged" />
                        </ItemTemplate>
                    </asp:TemplateField>
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
                    <asp:CommandField ButtonType="Image" CausesValidation="false" ShowEditButton="True" ShowDeleteButton="True" CancelImageUrl="~/cancel.png" UpdateImageUrl="~/accept.png" EditImageUrl="~/edit.png" DeleteImageUrl="~/delete.png" />
                </Columns>
            </asp:GridView>
        </div>

        <asp:LinkButton ID="btnDelete" CausesValidation="False" runat="server" OnClick="btnDelete_Click" CssClass="btn btn-danger">
            <img src="deleteall.png" width="20" />
            &nbsp;Delete selected
        </asp:LinkButton>

        <asp:LinkButton ID="ResetButton" CausesValidation="False" runat="server" OnClick="btnReset_Click" CssClass="btn btn-secondary"><i class="bi bi-trash"></i>&nbsp;Reset</asp:LinkButton>
        <asp:LinkButton ID="HideButton" CausesValidation="False" runat="server" OnClick="btnHide_Click" CssClass="btn btn-secondary"><i class="bi bi-eye-slash"></i>&nbsp;Hide</asp:LinkButton>

    </asp:Panel>
    <br>
    <asp:LinkButton ID="ShowButton" CausesValidation="False" runat="server" OnClick="btnShow_Click" CssClass="btn btn-light"><i class="bi bi-eye"></i>&nbsp;Show</asp:LinkButton>

    <div class="jumbotron">
        <div class="form-group">
            <p>Export data from database in an Excel in runtime</p>
            <asp:LinkButton ID="ExportButton" CausesValidation="false" runat="server" tagname="ExportButt" OnClick="btnExport_Click" OnClientClick="showExportSpinner();" CssClass="btn btn-primary"><i class="bi bi-file-earmark-excel"></i>&nbsp;Export</asp:LinkButton>
            <br>
        </div>
        <br>
    </div>

</asp:Content>
