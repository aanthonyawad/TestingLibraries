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
            FileStream docStream = new FileStream(wwwPath+ "/SyndicationInterestCertificate.docx", FileMode.Open, FileAccess.ReadWrite,FileShare.ReadWrite);

            WordDocument wordDocument = new WordDocument(docStream, Syncfusion.DocIO.FormatType.Automatic);
            
            
            
            // search for maturity Date to replace
            this.changeText("{maturityDateData1}", DateTime.UtcNow.ToLongDateString(),wordDocument);
            this.changeText("{maturityDateData2}", DateTime.UtcNow.ToLongDateString(), wordDocument);
            this.changeText("{maturityDateData3}", DateTime.UtcNow.ToLongDateString(), wordDocument);
            // search for maturity Day to replace
            this.changeText("{maturityDayData}", "" + DateTime.UtcNow.Day, wordDocument);
            // search for borrower name to replace
            this.changeText("{borrowerNameData}", "Jhon Borrower", wordDocument);
            // search for mortgage amount to replace
            this.changeText("{mortgageAmountData}", "$250,000", wordDocument);
            // search for property address to replace
            this.changeText("{propertyAddressData}", "123 Centre Street., Toronto, ON", wordDocument);
            // search for pid to replace
            this.changeText("{pidData}", "123-456-789", wordDocument);
            // search for crn to replace
            this.changeText("{crnData}", "987-654-321", wordDocument);
            // search for interest rate to replace
            this.changeText("{interestRateData}", "6%", wordDocument);
            // search for total number of units to replace
            this.changeText("{totalNumberOfUnitsData}", "5", wordDocument);
            // search for units invested to replace
            this.changeText("{unitsInvestedData}", "1", wordDocument);
            // search for percentage of invested data to replace
            this.changeText("{percentageInvestedData}", "20%", wordDocument);
            // search for trustee name data to replace
            this.changeText("{trusteeNameData}", "The trustee name company", wordDocument);
            // search for trustee name data to replace
            this.changeText("{totalUnitsAmountData}", "$50,000", wordDocument);


            //get the paragraph from the parent word document and add the signature to it
            TextSelection adminSignature = wordDocument.Find("{sig_trustee}", false, false);
            WTextRange signature = adminSignature.GetAsOneRange();
            signature.Text = "";
            //Adds image to  the paragraph
            FileStream imageStream = new FileStream(wwwPath +"/signature_2.png", FileMode.Open, FileAccess.Read);
            IWPicture picture = signature.OwnerParagraph.AppendPicture(imageStream);
            //Sets height and width for the image
            picture.Height = 75;
            picture.Width = 200;


            //opens the doc for rendeering 
            DocIORenderer render = new DocIORenderer();
            render.Settings.ChartRenderingOptions.ImageFormat = ExportImageFormat.Jpeg;
            PdfDocument pdfDocument = render.ConvertToPDF(wordDocument);
            render.Dispose();
            wordDocument.Dispose();

            //Saves the PDF file
            MemoryStream outputStream = new MemoryStream();
            pdfDocument.Save(outputStream);
            using (FileStream pdfStream = System.IO.File.Create(@"c:\temp\syndicationInterest.pdf"))
            {
                pdfDocument.Save(pdfStream);
            }
           
           
            //Closes the instance of PDF document object
            pdfDocument.Close();
            return "converted Word file to pdf...";
        }

        public void changeText(string key,string value,WordDocument wordDocument )
        {
            TextSelection keySelection = wordDocument.Find(key, false, false);
            if (keySelection == null) 
            {
                return;
            }
            WTextRange keyTextRange = keySelection.GetAsOneRange();
            keyTextRange.Text = value;
        } 

    }
}