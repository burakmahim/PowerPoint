using System;
using System.IO;
using System.Text;
using System.Web.Mvc;
using Microsoft.SqlServer.Server;
using PowerPointLibrary;
using PowerPointLibrary.Services;

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
                // BASE64 stringe dönüştür
                string base64Pdf = Convert.ToBase64String(pdfBytes);

                // ViewBag ile View'a gönder
                ViewBag.PowerPointPdf = "data:application/pdf;base64," + base64Pdf;
                ViewBag.XmlContent = xmlContent;
                return View("Index");
            }
            catch (Exception ex)
            {
                ViewBag.Error = "PDF oluşturulamadı: " + ex.Message;
                return View("Index");
            }
        }

        [HttpPost]
        [ValidateInput(false)]
        public ActionResult GenerateExcelPdf(string xmlContent)
        {
            try
            {
                byte[] pdfBytes = ExcelConverter.ConvertToPdf(xmlContent);

                // BASE64 stringe dönüştür
                string base64Pdf = Convert.ToBase64String(pdfBytes);

                // ViewBag ile View'a gönder
                ViewBag.ExcelPdf = "data:application/pdf;base64," + base64Pdf;
                ViewBag.XmlContent = xmlContent;
                return View("Index");

            }
            catch (Exception ex)
            {
                ViewBag.Error = "PDF oluşturulamadı: " + ex.Message;
                return View("Index");
            }
        }



    }
}