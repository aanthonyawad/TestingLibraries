using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using IronPdf;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;

namespace TestingLibraries.Controllers
{
    public class IronPdfController : Controller
    {
        private IHostingEnvironment Environment;

        public IronPdfController(IHostingEnvironment _environment)
        {
            Environment = _environment;
        }
        public string Index()
        {
            var Renderer = new IronPdf.HtmlToPdf();
            string wwwPath = this.Environment.WebRootPath;
            string contentPath = this.Environment.ContentRootPath;
            

            var PDF = PdfDocument.FromFile(wwwPath+ "/MortgageAdministrationAgreementOriginal.pdf");
            var date1 = PDF.Form.GetFieldByName("date1");
            date1.Value = DateTime.UtcNow.ToLongDateString();
            date1.SetFont(IronPdf.Forms.Enums.FontTypes.HelveticaBold);
            date1.ReadOnly = true;

            var date2 = PDF.Form.GetFieldByName("date2");
            date2.Value = DateTime.UtcNow.ToLongDateString()+",";
            date2.ReadOnly = true;


            var date3 = PDF.Form.GetFieldByName("date3");
            date3.Value = DateTime.UtcNow.Month + ","+ DateTime.UtcNow.Day +"," +DateTime.UtcNow.Year;
            date3.ReadOnly = true;


            PDF.RemovePage(PDF.PageCount - 1);
            PDF.AppendPdf(Renderer.RenderHtmlAsPdf(addSignature(wwwPath+"/signature_2.png")));
            PDF.SaveAs(@"c:\temp\irongenerated.pdf");
            return "generated with iron pdf";
        }

        private string addSignature(string image)
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendFormat("<html><head><style>");
            sb.Append("p{margin-bottom: 4px;}hr.solid-3{border-top: 3px solid black;}hr.solid-1{border-top: 1px solid black;}.text-muted{color: rgba(0, 0, 0, 0.54) !important;}.primary-color{color: #1976d2;}.container{margin-left: 20px !important; margin-left: 20px !important; -webkit-box-sizing: border-box; box-sizing: border-box;}.data-title-container{-webkit-box-orient: vertical; -webkit-box-direction: normal; -ms-flex-direction: column; flex-direction: column; -ms-flex-pack: distribute; justify-content: space-around; -webkit-box-align: start; -ms-flex-align: start; align-items: flex-start; display: -webkit-box; display: -ms-flexbox; display: flex; -ms-flex-wrap: nowrap; flex-wrap: nowrap;}.data-sub-title-container{-webkit-box-orient: vertical; -webkit-box-direction: normal; -ms-flex-direction: column; flex-direction: column; -ms-flex-pack: distribute; justify-content: space-around; -webkit-box-align: start; -ms-flex-align: start; align-items: flex-start; display: -webkit-box; display: -ms-flexbox; display: flex; -ms-flex-wrap: nowrap; flex-wrap: nowrap;}.data-center-center{-webkit-box-orient: vertical; -webkit-box-direction: normal; -ms-flex-direction: column; flex-direction: column; -ms-flex-pack: distribute; justify-content: space-around; -webkit-box-align: center; -ms-flex-align: center; align-items: center; display: -webkit-box; display: -ms-flexbox; display: flex; -ms-flex-wrap: nowrap; flex-wrap: nowrap;}.data-section-title-container{-webkit-box-orient: horizontal; -webkit-box-direction: normal; -ms-flex-direction: row; flex-direction: row; -webkit-box-pack: start; -ms-flex-pack: start; justify-content: flex-start; -webkit-box-align: start; -ms-flex-align: start; align-items: flex-start; display: -webkit-box; display: -ms-flexbox; display: flex; -ms-flex-wrap: nowrap; flex-wrap: nowrap;}.data-section-info-column{-webkit-box-orient: vertical; -webkit-box-direction: normal; -ms-flex-direction: column; flex-direction: column; -webkit-box-pack: center; -ms-flex-pack: center; justify-content: center; -webkit-box-align: center; -ms-flex-align: center; align-items: center; display: -webkit-box; display: -ms-flexbox; display: flex; -ms-flex-wrap: nowrap; flex-wrap: nowrap;}.data-section-info-container{min-width: 300px; -webkit-box-orient: horizontal; -webkit-box-direction: normal; -ms-flex-direction: row; flex-direction: row; -webkit-box-pack: start; -ms-flex-pack: start; justify-content: flex-start; -webkit-box-align: baseline; -ms-flex-align: baseline; align-items: baseline; display: -webkit-box; display: -ms-flexbox; display: flex; -ms-flex-wrap: nowrap; flex-wrap: nowrap;}.section-item-grow{-webkit-box-flex: 2; -ms-flex-positive: 2; flex-grow: 2;}.flex-line-container{-webkit-box-orient: vertical; -webkit-box-direction: normal; -ms-flex-direction: column; flex-direction: column; -webkit-box-pack: none; -ms-flex-pack: none; justify-content: none; -webkit-box-align: none; -ms-flex-align: none; align-items: none; display: -webkit-box; display: -ms-flexbox; display: flex;}.flex-line-container img{width: 300px; height: 100px;}.fs-10{font-size: 10px !important;}.fs-12{font-size: 12px !important;}.fs-14{font-size: 14px !important;}.fs-16{font-size: 16px !important;}.fs-18{font-size: 18px !important;}.fs-20{font-size: 20px !important;}.m-0{margin: 0 !important;}.mt-0,.my-0{margin-top: 0 !important;}.mb-0,.my-0{margin-bottom: 0 !important;}.mr-4,.mx-4{margin-right: 1.5rem !important;}.ml-4,.mx-4{margin-left: 1.5rem !important;}.mt-4,.my-4{margin-top: 1.5rem !important;}.mb-4,.my-4{margin-bottom: 1.5rem !important;}.m-3{margin: 1rem !important;}.mt-3,.my-3{margin-top: 1rem !important;}.mr-3,.mx-3{margin-right: 1rem !important;}.mb-3,.my-3{margin-bottom: 1rem !important;}");
            sb.AppendFormat("</style></head><body>");
            sb.AppendFormat("<div ><div class=\"data-section-title-container mb-4\">");
            sb.AppendFormat("<p><b>IN WITNESS WHEREOF</b>&nbsp;the Parties have duly executed this Agreement as of the date first written above.</p>");
            sb.AppendFormat("</div>");
            
            
            //ADMIN SIG
            sb.AppendFormat("<div class=\"data-section-info-column\">");
            sb.AppendFormat("<div class=\"data-section-info-container mb-4\">");
            sb.AppendFormat("<div>");
            sb.AppendFormat("<p><b>MM ADMINISTRATION INC.</b></p>");
            sb.AppendFormat("<p>Per:</p>");
            sb.AppendFormat("<div class=\"flex-line-container\">");
            sb.AppendFormat("<img src=\""+image+"\"></img>");
            sb.AppendFormat("<div>");
            sb.AppendFormat("<hr class=\"solid-1  mb-0\"></hr>");
            sb.AppendFormat("</div>");
            sb.AppendFormat("</div>");
            sb.AppendFormat("<p>Name:</p>");
            sb.AppendFormat("<p>Title:</p>");
            sb.AppendFormat("</div>");
            sb.AppendFormat("</div>");

            //CORP SIG
            sb.AppendFormat("<div class=\"data-section-info-column\">");
            sb.AppendFormat("<div class=\"data-section-info-container mb-4\">");
                sb.AppendFormat("<div>");
                sb.AppendFormat("<p><b>MM Corporation</b></p>");
                sb.AppendFormat("<p>Per:</p>");
                sb.AppendFormat("<div class=\"flex-line-container\">");
                sb.AppendFormat("<img src=\"" + image + "\"></img>");
                sb.AppendFormat("<div>");
                sb.AppendFormat("<hr class=\"solid-1  mb-0\"></hr>");
                sb.AppendFormat("</div>");
                sb.AppendFormat("</div>");
                sb.AppendFormat("<p>Name:</p>");
                sb.AppendFormat("<p>Title:</p>");
                sb.AppendFormat("</div>");

            sb.AppendFormat("</div>");

            sb.AppendFormat("</div><body></html>");
            return sb.ToString();
        }
    }
}