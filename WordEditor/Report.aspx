<%@ Page Language="C#" MasterPageFile="~/Site.Master" ValidateRequest="false" AutoEventWireup="true" CodeBehind="Report.aspx.cs" Inherits="WordEditor.Report" %>



<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <script type="text/javascript" src="http://ajax.googleapis.com/ajax/libs/jquery/1.7.2/jquery.min.js"></script>
    <script src="http://ajax.aspnetcdn.com/ajax/jquery.ui/1.8.9/jquery-ui.js" type="text/javascript"></script>
    <link href="http://ajax.aspnetcdn.com/ajax/jquery.ui/1.8.9/themes/blitzer/jquery-ui.css" rel="stylesheet" type="text/css" />
    
    <style>

    body{
        padding-top:100px;
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

<form>
    <div class="grid-container">

        <div class="grid-child purple">
            <textarea id="dataToCopy" placeholder="Write Here." runat="server"></textarea>
            <button id="Create" runat="server">Create Report</button>
            <button id="Download" runat="server" >Download Report</button>

            <asp:DropDownList ID="DropDownList" runat="server"
            onselectedindexchanged="DropDownList1_SelectedIndexChanged" AutoPostBack=true></asp:DropDownList>
            <br />
            <button id="Open" runat="server">Open Report</button>

            <asp:Button ID="SaveFile" style="width:100px" runat="server" Text="Save" OnClick="SaveReport" Visible="true" />
        </div>

        <div class="grid-child green">  
            <asp:TextBox ID="textbox1"  runat="server" ></asp:TextBox>
        </div>

    </div>
</form>

    <asp:Panel runat="server" ID="pnlEditor">
    <script type="text/javascript">
        tinymce.init({
            selector: '#<%=textbox1.ClientID%>',
            plugins: 'anchor autolink charmap codesample emoticons image link lists media searchreplace table visualblocks wordcount checklist mediaembed casechange export formatpainter pageembed linkchecker a11ychecker tinymcespellchecker permanentpen powerpaste advtable advcode editimage tinycomments tableofcontents footnotes mergetags autocorrect typography inlinecss ExportToDoc',
            toolbar: 'undo redo | blocks fontfamily fontsize | bold italic underline strikethrough | link image media table mergetags | addcomment showcomments | spellcheckdialog a11ycheck typography | align lineheight | checklist numlist bullist indent outdent | emoticons charmap | removeformat | ExportToDoc',
            tinycomments_mode: 'embedded',
            tinycomments_author: 'Author name',
            branding: false,
            mergetags_list: [
                { value: 'First.Name', title: 'First Name' },
                { value: 'Email', title: 'Email' },
            ]
        });
    </script>
</asp:Panel>
    
    <script>

        // Get from textarea and set content of RTE
        { // Get from textarea and set content of RTE
            document.getElementById('<%=Create.ClientID%>').addEventListener("click", copyToClipBoard);

            function copyToClipBoard() {
                var textarea = document.getElementById('<%=dataToCopy.ClientID%>').value;
                Setcontent(textarea);
            }

            function Setcontent(textarea) {
                var ContentSet = tinymce.activeEditor.setContent(textarea);
                event.preventDefault();
            }

        }

        // Download RTE Content
        {
            document.getElementById('<%=Download.ClientID%>').addEventListener("click", function () { Export2Word() });

            function Export2Word(filename = 'content') {
                var preHtml = "<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'><head><meta charset='utf-8'><title>Export HTML To Doc</title></head><body>";
                var postHtml = "</body></html>";
                var html = preHtml + tinymce.activeEditor.getContent() + postHtml;
                //alert(html);

                var blob = new Blob(['\ufeff', html], {
                    type: 'application/msword'
                });

                // Specify link url
                var url = 'data:application/vnd.ms-word;charset=utf-8,' + encodeURIComponent(html);

                // Specify file name
                filename = filename ? filename + '.doc' : 'document.doc';

                // Create download link element
                var downloadLink = document.createElement("a");

                document.body.appendChild(downloadLink);

                if (navigator.msSaveOrOpenBlob) {
                    navigator.msSaveOrOpenBlob(blob, filename);
                } else {
                    // Create a link to the file
                    downloadLink.href = url;

                    // Setting the file name
                    downloadLink.download = filename;

                    //triggering the function
                    downloadLink.click();
                }

                document.body.removeChild(downloadLink);
            }

        }


    </script>
</asp:Content>

