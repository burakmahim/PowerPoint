using System;
using System.IO;
using System.Text;
using System.Web.Mvc;
using PowerPointLibrary;

namespace PowerPointApp.Mvc.Controllers
{
    public class PresentationController : Controller
    {
        // GET: /Presentation/
        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]
        [ValidateInput(false)] // <--- Zararlı olarak algılanabilecek karakterleri kabul et
        public ActionResult DownloadPptx(string xmlContent)
        {
            try
            {
                byte[] pptBytes = PowerPointGenerator.CreatePresentationFromXml(xmlContent);
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
        public ActionResult GeneratePptx(string xmlContent)
        {
            if (string.IsNullOrWhiteSpace(xmlContent))
                return new HttpStatusCodeResult(400, "XML içeriği boş olamaz.");

            try
            {
                byte[] pptBytes = PowerPointGenerator.CreatePresentationFromXml(xmlContent);
                return File(pptBytes,
                            "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                            "presentation.pptx");
            }
            catch (Exception ex)
            {
                return new HttpStatusCodeResult(500, $"Sunum oluşturulamadı: {ex.Message}");
            }
        }

        [HttpPost]
        public ActionResult GeneratePdf(string xmlContent)
        {
            if (string.IsNullOrWhiteSpace(xmlContent))
                return new HttpStatusCodeResult(400, "XML içeriği boş olamaz.");

            try
            {
                byte[] pdfBytes = PowerPointGenerator.ConvertToPdf(xmlContent);
                return File(pdfBytes, "application/pdf", "presentation.pdf");
            }
            catch (Exception ex)
            {
                return new HttpStatusCodeResult(500, $"PDF dönüştürme başarısız: {ex.Message}");
            }
        }

        [HttpPost]
        [ValidateInput(false)] // 🛡 Bu satır HTML/XML içeriği kabul eder
        public ActionResult Generate(string xmlContent, string format)
        {
            if (string.IsNullOrWhiteSpace(xmlContent))
                return Content("XML içeriği boş olamaz.");

            byte[] result;
            string mime, filename;

            if (format == "pdf")
            {
                result = PowerPointGenerator.ConvertToPdf(xmlContent);
                mime = "application/pdf";
                filename = "sunum.pdf";
            }
            else
            {
                result = PowerPointGenerator.CreatePresentationFromXml(xmlContent);
                mime = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
                filename = "sunum.pptx";
            }

            return File(result, mime, filename);
        }

    }
}
