using System;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;

namespace TestingLibraries.Controllers
{
    public class HelloWorldController : Controller
    {
        private IHostingEnvironment Environment;

        public HelloWorldController(IHostingEnvironment _environment)
        {
            Environment = _environment;
        }

        // 
        // GET: /HelloWorld/

        public string Index()
        {
            string wwwPath = this.Environment.WebRootPath;
            string contentPath = this.Environment.ContentRootPath;
            // load the DOC file to be converted
            var document = new Aspose.Words.Document(wwwPath+"/SyndicationInterestCertificate.docx");
            Aspose.Words.DocumentBuilder builder = new Aspose.Words.DocumentBuilder(document);
            var findText = @"{maturityDate1}";
            document.Range.Replace(findText, DateTime.Now.ToLongDateString(), new Aspose.Words.Replacing.FindReplaceOptions());
            string signature = "{sig_trustee}";
            //builder.InsertImage(wwwPath + "/signature_2.png",
            //                Aspose.Words.Drawing.RelativeHorizontalPosition.Margin,
            //                100,
            //                Aspose.Words.Drawing.RelativeVerticalPosition.Margin,
            //                100,
            //                200,
            //                100,
            //                Aspose.Words.Drawing.WrapType.Square);

            int paragraphId = 0;
            Aspose.Words.NodeCollection paragraphs = document.GetChildNodes(Aspose.Words.NodeType.Paragraph, true);
            for(int i = 0; i  < paragraphs.Count; i++)
            {   
                bool isSigParag = paragraphs[i].GetText().Contains(signature);
                if (isSigParag)
                {
                    builder.MoveTo(paragraphs[i]);
                    builder.InsertImage(wwwPath + "/signature_2.png",
                            Aspose.Words.Drawing.RelativeHorizontalPosition.Margin,
                            100,
                            Aspose.Words.Drawing.RelativeVerticalPosition.Margin,
                            0,
                            50,
                            100,
                            Aspose.Words.Drawing.WrapType.Square);
                }
            }
            // save DOC as a PDF
            document.Save(@"c:\temp\asposeInvestment.pdf", Aspose.Words.SaveFormat.Pdf);

            return "pdf generated using aspose";
        }
    }
}