using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.Office.Interop.Word;
using System.Text.RegularExpressions;
using System.Linq;
using System.Web;
using System.Reflection;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.Hosting;
using System.Runtime.InteropServices.ComTypes;
using System.Runtime.InteropServices;
using System.Text;
using System.Drawing;
using System.Windows.Forms;
using System.Web.UI.HtmlControls;
using Application = Microsoft.Office.Interop.Word.Application;
using Paragraph = Microsoft.Office.Interop.Word.Paragraph;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using OpenXmlPowerTools;
using HtmlAgilityPack;
using Xceed.Words.NET;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;
using Aspose.Words;

namespace WordEditor {
  public partial class Report : System.Web.UI.Page {


    public string wordContent = "";
    protected void Page_Load(object sender, EventArgs e) {
      if (!IsPostBack) {
        var reportFolderPath = HostingEnvironment.MapPath("~/Reports");
        IEnumerable<string> folders = Directory.GetFileSystemEntries(reportFolderPath, "*.doc");
        folders = Directory.GetFileSystemEntries(reportFolderPath, "*.doc");
        folders = folders.Select(o => Path.GetFileNameWithoutExtension(o)); ;
        //folders = folders.Select(o => Path.GetFileNameWithoutExtension(o));
        //folders = folders.Select(o => Path.GetFileName(o));
        DropDownList.DataSource = folders;
        DropDownList.DataBind();
      }
    }
    protected void DropDownList1_SelectedIndexChanged(object sender, EventArgs e) {
      //System.Diagnostics.Process.Start(DropdownList.SelectedValue);
      var s = DropDownList.SelectedValue;
      Upload(s);
    }
    protected void Upload(string s) {
      object documentFormat = 8;
      string randomName = s;
      object htmlFilePath = HttpContext.Current.Server.MapPath("~/Reports/") + randomName + ".htm";
      string directoryPath = HttpContext.Current.Server.MapPath("~/Temp/") + randomName + "_files";
      object fileSavePath = HttpContext.Current.Server.MapPath("~/Reports/") + randomName;

      //If Directory not present, create it.
      if (!Directory.Exists(HttpContext.Current.Server.MapPath("~/Reports/"))) {
        Directory.CreateDirectory(HttpContext.Current.Server.MapPath("~/Reports/"));
      }

      //Open the word document in background.
      Application applicationclass = new Application();
      applicationclass.Documents.Open(ref fileSavePath);
      applicationclass.Visible = false;
      Microsoft.Office.Interop.Word.Document document = applicationclass.ActiveDocument;

      //Save the word document as HTML file.
      document.SaveAs(ref htmlFilePath, ref documentFormat);

      //Close the word document.
      document.Close();

      //Read the saved Html File.
      //string[] byteWord = System.IO.File.read(htmlFilePath.ToString());
      string wordHTML = System.IO.File.ReadAllText(htmlFilePath.ToString());
      //Loop and replace the Image Path.
      foreach (Match match in Regex.Matches(wordHTML, "<v:imagedata.+?src=[\"'](.+?)[\"'].*?>", RegexOptions.IgnoreCase)) {
        wordHTML = Regex.Replace(wordHTML, match.Groups[1].Value, "Reports/" + match.Groups[1].Value);
      }

      System.IO.File.Delete(htmlFilePath.ToString());
      System.IO.File.Delete(fileSavePath.ToString());
      Directory.Delete(fileSavePath.ToString() + "_files", true);

      //dvWord.InnerHtml = wordHTML;

      textbox1.Text = wordHTML;
    }


    public void SaveReport(object sender, EventArgs e) {
      //Response.Write("helllloo");
      var filename = DropDownList.SelectedValue;
      object fileSavePath = HttpContext.Current.Server.MapPath("~/Reports/") + filename+".doc";
      object htmlFilePath = HttpContext.Current.Server.MapPath("~/Reports/temp.html");

      /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

      Aspose.Words.Document document = new Aspose.Words.Document();

      // Create a document builder
      Aspose.Words.DocumentBuilder builder = new Aspose.Words.DocumentBuilder(document);

      // Insert HTML
      builder.InsertHtml(textbox1.Text);
      //builder.InsertHtml(textbox.Text);

      // Save as DOCX
      document.Save(fileSavePath.ToString(), SaveFormat.Doc);

      

      Response.Redirect(Request.Url.ToString());
    }







  }



  }
