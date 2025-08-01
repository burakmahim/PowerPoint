using System;
using System.Data;
using System.IO;
using System.Text;
using System.Web.Mvc;
using System.Xml.Linq;
using Microsoft.SqlServer.Server;
using PowerPointLibrary;


namespace PowerPointApp.Mvc.Controllers
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
        public ActionResult GenerateExcelFromXml(string xmlContent)
        {
            if (string.IsNullOrWhiteSpace(xmlContent))
                return new HttpStatusCodeResult(400, "XML boş olamaz.");

            try
            {
                byte[] result = ExcelLibrary.CreateExcelFromCustomXml(xmlContent);
                return File(result, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "veriler.xlsx");
            }
            catch (Exception ex)
            {
                return new HttpStatusCodeResult(500, $"Excel oluşturulamadı: {ex.Message}");
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
                byte[] pdfBytes = PowerPointGenerator.ConvertToPdf(xmlContent);
                return File(pdfBytes, "application/pdf");
            }
            catch (Exception ex)
            {
                ViewBag.Error = ex.Message;
                return View("Index");
            }
        }

        [HttpPost]
        [ValidateInput(false)]
        public ActionResult GenerateExcelPdf(string xmlContent)
        {

            if (string.IsNullOrWhiteSpace(xmlContent))
            {
                ViewBag.Error = "XML içeriği boş gönderildi.";
                return View("Index");
            }

            try
            {
                byte[] pdfBytes = ExcelLibrary.ConvertToPdf(xmlContent);
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
