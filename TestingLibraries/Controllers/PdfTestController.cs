using System;
using System.IO;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.OfficeChart;
using Syncfusion.Pdf;


namespace TestingLibraries.Controllers
{
    public class PdfTestController : Controller
    {
        private IHostingEnvironment Environment;

        public PdfTestController(IHostingEnvironment _environment)
        {
            Environment = _environment;
        }
        // 
        // GET: /PdfTest/

        public string Index()
        {

            
            string wwwPath = this.Environment.WebRootPath;
            string contentPath = this.Environment.ContentRootPath;
            // get file and lock it for readandwrite
            FileStream docStream = new FileStream(wwwPath+"/document_3_template.docx", FileMode.Open, FileAccess.ReadWrite,FileShare.ReadWrite);

            WordDocument wordDocument = new WordDocument(docStream, Syncfusion.DocIO.FormatType.Automatic);
            // seach for word occurences to replace
            TextSelection textSelection = wordDocument.Find("{date}", false, false);

            WTextRange[] textRanges = textSelection.GetRanges();
            foreach (WTextRange textRange in textRanges)
            {
                textRange.Text = DateTime.UtcNow.ToLongDateString();
            }

            //opens the doc for rendeering 
            DocIORenderer render = new DocIORenderer();
            render.Settings.ChartRenderingOptions.ImageFormat = ExportImageFormat.Jpeg;
            PdfDocument pdfDocument = render.ConvertToPDF(wordDocument);
            render.Dispose();
            wordDocument.Dispose();

            //Saves the PDF file
            MemoryStream outputStream = new MemoryStream();
            pdfDocument.Save(outputStream);
            using (FileStream pdfStream = System.IO.File.Create(@"c:\temp\pdfGenerated.pdf"))
            {
                pdfDocument.Save(pdfStream);
            }
           
           
            //Closes the instance of PDF document object
            pdfDocument.Close();
            return "converted Word file to pdf...";
        }

    }
}