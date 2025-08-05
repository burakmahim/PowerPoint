using System;
using System.Data;
using System.IO;
using System.Text;
using System.Web.Mvc;
using System.Xml.Linq;
using Microsoft.SqlServer.Server;
using OfficeAppLibrary;


namespace OfficeApp.Mvc.Controllers
{
    public class PresentationController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]
        [ValidateInput(false)]
        public ActionResult DownloadPptx(string xmlContent)
        {
            try
            {
                byte[] pptBytes = PowerPointLibrary.PowerPointGenerator.CreatePresentationFromXml(xmlContent);
                return File(pptBytes,
                    "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    "Sunum.pptx");
            }
            catch (Exception ex)
            {
                return Content("Hata oluştu: " + ex.Message);
            }
        }

        [HttpPost]
        [ValidateInput(false)]
        public ActionResult ViewPdf(string xmlContent)
        {
            if (string.IsNullOrWhiteSpace(xmlContent))
            {
                ViewBag.Error = "XML içeriği boş gönderildi.";
                return View("Index");
            }

            try
            {
                byte[] pdfBytes = PowerPointLibrary.PowerPointGenerator.ConvertToPdf(xmlContent);
                return File(pdfBytes, "application/pdf");
            }
            catch (Exception ex)
            {
                ViewBag.Error = ex.Message;
                return View("Index");
            }
        }
    }
}
