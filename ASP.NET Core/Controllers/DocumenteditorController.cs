using DocumentEditorApp.Models;
using Microsoft.AspNetCore.Cors;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System;
using System.Data.OleDb;
using System.IO;
using System.Collections.Generic;
using EJ2DocumentEditor = Syncfusion.EJ2.DocumentEditor;
using Syncfusion.DocIORenderer;
using Syncfusion.Pdf;
using DocIOWordDocument = Syncfusion.DocIO.DLS.WordDocument;
using System.Data.SqlClient;
using System.Text;
using System.Data;

namespace DocumentEditorApp.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class DocumenteditorController : ControllerBase
    {
        private IWebHostEnvironment hostEnvironment;

        private string dataSourcePath;

        private string connectionString;
        public DocumenteditorController(IWebHostEnvironment environment)
        {
            this.hostEnvironment = environment;
            this.dataSourcePath = Path.Combine(this.hostEnvironment.ContentRootPath,"AppData"+"\\DocumentInfo.accdb");
            this.connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + this.dataSourcePath + ";User Id=admin;Password=;";
        }

        [Route("Import")]
        public string Import(IFormCollection data)
        {  
        if (data.Files.Count == 0)
        return null;
        System.IO.Stream stream = new System.IO.MemoryStream();
        Microsoft.AspNetCore.Http.IFormFile file = data.Files[0];
        int index = file.FileName.LastIndexOf('.');
        string type = index > -1 && index < file.FileName.Length - 1 ?
        file.FileName.Substring(index) : ".docx";
        file.CopyTo(stream);
        stream.Position = 0;
        
        EJ2DocumentEditor.WordDocument document = EJ2DocumentEditor.WordDocument.Load(stream, GetFormatType(type.ToLower()));
        string json = Newtonsoft.Json.JsonConvert.SerializeObject(document);
        document.Dispose();
        return json;
        
        }
        public class CustomClipboarParameter
        {
            public string content { get; set; }
            public string type { get; set; }
        }


        [Route("SystemClipboard")]
        public string SystemClipboard([FromBody]CustomClipboarParameter param)
        {
            if (param.content != null && param.content != "")
            {
                try
                {
                    Syncfusion.EJ2.DocumentEditor.WordDocument document = Syncfusion.EJ2.DocumentEditor.WordDocument.LoadString(param.content, GetFormatType(param.type.ToLower()));
                    string json = Newtonsoft.Json.JsonConvert.SerializeObject(document);
                    document.Dispose();
                    return json;
                }
                catch (Exception)
                {
                    return "";
                }
            }
            return "";
        }

        public class CustomRestrictParameter
        {
            public string passwordBase64 { get; set; }
            public string saltBase64 { get; set; }
            public int spinCount { get; set; }
        }

        [Route("RestrictEditing")]
        public string[] RestrictEditing([FromBody]CustomRestrictParameter param)
        {
            if (param.passwordBase64 == "" && param.passwordBase64 == null)
                return null;
            return Syncfusion.EJ2.DocumentEditor.WordDocument.ComputeHash(param.passwordBase64, param.saltBase64, param.spinCount);
        }

        
        [Route("ImportFile")]
        public string ImportFile([FromBody]Documentdetails param)
        {
            string path = this.hostEnvironment.WebRootPath + "\\Files\\" + param.FileName; 
            try
            {
                string query = "SELECT EditorName FROM DocumentInfo WHERE DocumentName = '"+param.FileName+"'";
                string dataSourcePath = Path.Combine(hostEnvironment.ContentRootPath,"AppData"+"\\DocumentInfo.accdb");
                string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + dataSourcePath + ";User Id=admin;Password=;";
                DataTable table = getDatabaseData(connectionString,query);
                string Result = Newtonsoft.Json.JsonConvert.SerializeObject(table);
                string status = "";
                if (table.Rows[0][0] is string)
                    status = (string)table.Rows[0][0];
                if (status != ""){
                    return Result;
                }
                DateTime date = DateTime.Now;
                query = "UPDATE DocumentInfo SET EditorName = '"+param.userName+"', LastModifiedBy= '"+param.userName+"', LastModifiedTime =#"+date+"#  WHERE DocumentName = '"+param.FileName+"'";
                performCRUD(connectionString,query);
                Stream stream = System.IO.File.Open(path, FileMode.Open, FileAccess.ReadWrite);
                Syncfusion.EJ2.DocumentEditor.WordDocument document = Syncfusion.EJ2.DocumentEditor.WordDocument.Load(stream, GetFormatType(path));
                string json = Newtonsoft.Json.JsonConvert.SerializeObject(document);
                document.Dispose();
                stream.Dispose();
                return json;
            }
            catch(Exception ex)
            {
                return ex.Message;
            }
        }

        [Route("LogOut")]
        public void LogOut([FromBody]CustomParams param){
            string SQLquery = "UPDATE DocumentInfo SET EditorName = '' WHERE DocumentName = '"+ param.fileName +"'";
            performCRUD(connectionString,SQLquery);
        }
        
        [AcceptVerbs("Post")]
        [HttpPost]
        [Route("ExportPdf")]
        public FileStreamResult ExportPdf([FromBody] SaveParameter data)
        {
            // Converts the sfdt to stream
            Stream document = EJ2DocumentEditor.WordDocument.Save(data.Content, EJ2DocumentEditor.FormatType.Docx);
            Syncfusion.DocIO.DLS.WordDocument doc = new Syncfusion.DocIO.DLS.WordDocument(document, Syncfusion.DocIO.FormatType.Docx);
            //Instantiation of DocIORenderer for Word to PDF conversion 
            DocIORenderer render = new DocIORenderer();
            //Converts Word document into PDF document 
            PdfDocument pdfDocument = render.ConvertToPDF(doc);
            Stream stream = new MemoryStream();
            
            //Saves the PDF file
            pdfDocument.Save(stream);
            stream.Position = 0;
            pdfDocument.Close();         
            document.Close();
            return new FileStreamResult(stream, "application/pdf")
            {
                FileDownloadName = data.FileName
            };
        }
        public class SaveParameter
        {
            public string Content { get; set; }
            public string FileName { get; set; }
            public string UserName { get; set; }
        }

        
        [Route("ExportSFDT")]
        public void ExportSFDT([FromBody]Saveparameter data)
        {
            string name = data.FileName;
            string path = this.hostEnvironment.WebRootPath + "\\Files\\" + data.FileName;
            EJ2DocumentEditor.FormatType format = GetFormatType(data.FileName);
            if (string.IsNullOrEmpty(name))
            {
                name = "Document1.doc";
            }
            DocIOWordDocument document = Syncfusion.EJ2.DocumentEditor.WordDocument.Save(data.content);
            FileStream fileStream = new FileStream(path, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            document.Save(fileStream, GetDocIOFomatType(format));
            document.Close();
            fileStream.Close();
        }

        [Route("InsertRow")]

        public int InsertRow([FromBody]Documentdetails param)
        {

            string SQLquery = "SELECT COUNT(DocumentName) FROM DocumentInfo WHERE DocumentName = '"+param.FileName+"';";
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                OleDbCommand command = new OleDbCommand(SQLquery);

                command.Connection = connection;

                try
                {
                    connection.Open();
                    int count = (int)command.ExecuteScalar();
                    connection.Close();
                    if(count != 0){
                        return 0;
                    }
                    else {
                        DateTime date = DateTime.Now;

                        string insertSQL = "INSERT INTO DocumentInfo (DocumentName, AuthorName, LastModifiedTime, LastModifiedBy, EditorName) " +
                        "VALUES ('"+param.FileName+"', '"+param.userName+"', #"+date+"#, '"+param.userName+"', '"+param.userName+"')";
                
                        performCRUD(connectionString,insertSQL);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
            
            
            return 1;
        }

        [Route("RetriveDataSource")]

        public string RetriveDataSource()
        {
            
            var JSONString = "";
            string dataSourcePath = Path.Combine(hostEnvironment.ContentRootPath,"AppData"+"\\DocumentInfo.accdb");
            string SQLquery = "SELECT DocumentName,AuthorName,LastModifiedTime,LastModifiedBy FROM DocumentInfo";
            string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + dataSourcePath + ";User Id=admin;Password=;";
            DataTable table = getDatabaseData(connectionString,SQLquery);
            JSONString = Newtonsoft.Json.JsonConvert.SerializeObject(table);
            return JSONString;
        }

        [Route("DeleteRecords")]

        public void DeleteRecords([FromBody]CustomParams param){

            string SQLquery = "DELETE FROM DocumentInfo WHERE DocumentName = '"+param.fileName+"'";
            performCRUD(connectionString,SQLquery);
        }
        public void performCRUD(string connectionString, string SQLquery){

            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                OleDbCommand command = new OleDbCommand(SQLquery);

                command.Connection = connection;

                try
                {
                    connection.Open();
                    command.ExecuteNonQuery();
                    connection.Close();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
        }

        public DataTable getDatabaseData(string connectionString, string SQLquery){

            DataTable table = new DataTable();
            try {
                using(OleDbConnection con = new OleDbConnection(connectionString))
                {
                    using(OleDbCommand command = new OleDbCommand(SQLquery,con))
                    {
                        con.Open();
                        OleDbDataReader reader = command.ExecuteReader();

                        
                        table.Load(reader);
                        con.Close();
                    }
                }
            }
            catch (Exception ex){
                Console.WriteLine(ex.Message);
            }
            return table;
        }

        public class Documentdetails
        {
            public string userName { get; set; }
            public string FileName { get; set; }

        }
        public class Saveparameter
        {
            public string content { get; set; }
            public string FileName { get; set; }

        }

        //Save document in web server.
        [Route("Save")]
        public string Save([FromBody]CustomParameter param)
        {
            string path = this.hostEnvironment.WebRootPath + "\\Files\\" + param.fileName;
            Byte[] byteArray = Convert.FromBase64String(param.documentData);
            Stream stream = new MemoryStream(byteArray);
            EJ2DocumentEditor.FormatType type = GetFormatType(path);
            try
            {
                FileStream fileStream = new FileStream(path, FileMode.OpenOrCreate, FileAccess.ReadWrite);

                if (type != EJ2DocumentEditor.FormatType.Docx)
                {
                    Syncfusion.DocIO.DLS.WordDocument document = new Syncfusion.DocIO.DLS.WordDocument(stream, Syncfusion.DocIO.FormatType.Docx);
                    document.Save(fileStream, GetDocIOFomatType(type));
                    document.Close();
                }
                else
                {
                    stream.Position = 0;
                    stream.CopyTo(fileStream);
                }
                stream.Dispose();
                fileStream.Dispose();
                return "Sucess";
            }
            catch
            {
                Console.WriteLine("err");
                return "Failure";
            }
        }

        internal static EJ2DocumentEditor.FormatType GetFormatType(string fileName)
        {
            int index = fileName.LastIndexOf('.');
            string format = index > -1 && index < fileName.Length - 1 ? fileName.Substring(index + 1) : "";

            if (string.IsNullOrEmpty(format))
                throw new NotSupportedException("EJ2 Document editor does not support this file format.");
            switch (format.ToLower())
            {
                case "dotx":
                case "docx":
                case "docm":
                case "dotm":
                    return EJ2DocumentEditor.FormatType.Docx;
                case "dot":
                case "doc":
                    return EJ2DocumentEditor.FormatType.Doc;
                case "rtf":
                    return EJ2DocumentEditor.FormatType.Rtf;
                case "txt":
                    return EJ2DocumentEditor.FormatType.Txt;
                case "xml":
                    return EJ2DocumentEditor.FormatType.WordML;
                default:
                    throw new NotSupportedException("EJ2 Document editor does not support this file format.");
            }
        }

        internal static Syncfusion.DocIO.FormatType GetDocIOFomatType(EJ2DocumentEditor.FormatType type)
        {
            switch (type)
            {
                case EJ2DocumentEditor.FormatType.Docx:
                    return FormatType.Docx;
                case EJ2DocumentEditor.FormatType.Doc:
                    return FormatType.Doc;
                case EJ2DocumentEditor.FormatType.Rtf:
                    return FormatType.Rtf;
                case EJ2DocumentEditor.FormatType.Txt:
                    return FormatType.Txt;
                case EJ2DocumentEditor.FormatType.WordML:
                    return FormatType.WordML;
                default:
                    throw new NotSupportedException("DocIO does not support this file format.");
            }
        }

    }

        
}
