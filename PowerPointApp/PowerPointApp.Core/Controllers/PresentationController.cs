using Microsoft.AspNetCore.Mvc;
using System.Text;
using PowerPointLibrary;
using PowerPointLibrary.Exceptions;
using PowerPointLibrary.ExcelServices;


namespace PowerPointApp.Core.Controllers
{
    [Route("[controller]/[action]")]
    public class PresentationController : Controller
    {
        [HttpGet]
        public IActionResult Index()
        {
            return View();
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult DownloadPptx(string xmlContent)
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
        [ValidateAntiForgeryToken]
        public IActionResult GenerateExcelFromXml(string xmlContent)
        {
            if (string.IsNullOrWhiteSpace(xmlContent))
                return BadRequest("XML boş olamaz.");

            try
            {
                byte[] result = ExcelLibrary.CreateExcelFromCustomXml(xmlContent);
                return File(result, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "veriler.xlsx");
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Excel oluşturulamadı: {ex.Message}");
            }
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult ViewPdf(string xmlContent)
        {
            if (string.IsNullOrWhiteSpace(xmlContent))
            {
                ViewBag.Error = "XML içeriği boş gönderildi.";
                return View("Index");
            }

            try
            {
                byte[] pdfBytes = PowerPointGenerator.ConvertToPdf(xmlContent);
                string base64Pdf = Convert.ToBase64String(pdfBytes);

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
        [ValidateAntiForgeryToken]
        public IActionResult GenerateExcelPdf(string xmlContent)
        {
            try
            {
                byte[] pdfBytes = ExcelConverter.ConvertToPdf(xmlContent);
                string base64Pdf = Convert.ToBase64String(pdfBytes);

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