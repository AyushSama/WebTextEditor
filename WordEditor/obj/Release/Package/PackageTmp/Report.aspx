<%@ Page Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Report.aspx.cs" Inherits="WordEditor.Report" %>


<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">

    <script>
        
    </script>

    <style>
    body{
        padding-top:100px;
    }
    #FinalContent{
        display:none;
    }
    #dataToCopy{
        height: 40px;
        width:100%;
    }
    .grid-container {
        display: grid;
        grid-template-columns: 1fr 3fr;
    }
    button{
        width:150px;
    }
    #dropdown{
        width:150px;
        height:40px;
        margin-top:10px;
        margin-bottom: 10px;
    }
</style>

<form method="post">
    <div class="grid-container">

        <div class="grid-child purple">

            <textarea id="dataToCopy" placeholder="Write Here."></textarea>
            <button id="Create">Create Report</button>
            <button id="Save" >Save Report</button>
            <select name="templates" id="dropdown">
            </select>
            <button id="Open">Open Report</button>
            <button id="Delete">Delete Report</button>
        </div>

        <div class="grid-child green">
            <textarea id="tiny" ></textarea>
        </div>

        <div id="FinalContent">
        </div>

    </div>
</form>
  <br />

    <asp:DropDownList ID="DropDownList" runat="server"
            onselectedindexchanged="DropDownList1_SelectedIndexChanged" AutoPostBack=true></asp:DropDownList>

    <br />
    <br />
    <br />
    <br />
    <div id="temporary" runat="server"></div>
    



</asp:Content>

