<%@ Page Title="Home Page" Language="C#" Async="true" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="WF_Import._Default" %>

<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">

    <script>
        function showPopup() {
            document.getElementById('popup').style.display = 'block';
        }

        function closePopup() {
            document.getElementById('popup').style.display = 'none';
        }

    </script>

    <h2>Welcome</h2>
    <div class="jumbotron">
        <h1>Excel file manager</h1>
        <div class="form-group">
            <p>Here you can upload an excel file to import the data into the database.</p>
            <asp:FileUpload ID="fileUpload" ClientIDMode="Static" CssClass="btn btn-default" accept="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, application/vnd.ms-excel" OnChange="fileUpload_Changed" runat="server" />
            <asp:RequiredFieldValidator ID="RequiredFieldValidator1" runat="server" ControlToValidate="fileUpload" ErrorMessage="Devi prima selezionare un file da importare!<br>" Display="Dynamic"></asp:RequiredFieldValidator>
            <br>
            <br>
            <asp:LinkButton ID="ImportButton" runat="server" tagname="ImportButt" OnClick="btnUpload_Click" OnClientClick="showSpinner();" CssClass="btn btn-primary"><i class="bi bi-upload"></i>&nbsp;Import</asp:LinkButton>
            <br>
            <br>
        </div>
    </div>

    <div class="jumbotron" visible="false" id="errorDiv" runat="server">
        <h2>Errors:</h2>
        <br>
    </div>

    <div id="popup" class="popup">
        <div style="background: linear-gradient(90deg, #D7E1EC 0%, #FFFFFF 100%);" class="popup-content">
            <span class="popup-close" onclick="closePopup()">&times;</span>
            <div class="popup-image">
                <img runat="server" id="popupImg" src="check1.png" />
            </div>
            <p runat="server" id="popupTextSucc" style="color: black;" class="popup-text"></p>
            <p runat="server" id="popupTextRows" style="color: #0275d8; font-size: 16px;" class="popup-text"></p>
            <p runat="server" id="popupTextErr" style="color: #d9534f; font-size: 16px; margin-bottom: 20px;" class="popup-text"></p>
            <div class="popup-button">
                <asp:Button ID="redButton" Text="Go to the View page" CausesValidation="false" OnClick="redView" CssClass="btn btn-primary" runat="server" />
            </div>
        </div>
    </div>

    <div id="spinner" style="display: none">
        <img src="spinner.gif" alt="Loading..." width="50" />&nbsp;<b>Loading...</b>
    </div>

</asp:Content>
