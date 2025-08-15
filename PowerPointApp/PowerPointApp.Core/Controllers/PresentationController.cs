using Microsoft.AspNetCore.Mvc;
using PowerPointLibrary;
using PowerPointLibrary.PowerpointComponents;
using PowerPointLibrary.ExcelComponents;
using System;

namespace PowerPointApp.Controllers
{
    public class PresentationController : Controller
    {
        [HttpGet]
        public IActionResult Index()
        {
            ViewBag.XmlContent = "";
            return View();
        }

        [HttpPost]
        public IActionResult DownloadPptx([FromForm] string xmlContent)
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
        public IActionResult ViewPdf([FromForm] string xmlContent)
        {
            if (string.IsNullOrWhiteSpace(xmlContent))
            {
                ViewBag.Error = "XML içeriği boş gönderildi.";
                ViewBag.XmlContent = "";
                return View("Index");
            }

            try
            {
                byte[] pdfBytes = PowerPointPdfConverter.ConvertToPdf(xmlContent);
                return File(pdfBytes, "application/pdf");
            }
            catch (Exception ex)
            {
                ViewBag.Error = ex.Message;
                ViewBag.XmlContent = xmlContent;
                return View("Index");
            }
        }

        [HttpPost]
        public IActionResult GenerateExcelFromXml([FromForm] string xmlContent)
        {
            if (string.IsNullOrWhiteSpace(xmlContent))
                return BadRequest("XML boş olamaz.");

            try
            {
                byte[] result = ExcelLibrary.CreateExcelFromCustomXml(xmlContent);
                return File(result,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    "veriler.xlsx");
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Excel oluşturulamadı: {ex.Message}");
            }
        }

        [HttpPost]
        public IActionResult GenerateExcelPdf([FromForm] string xmlContent)
        {
            if (string.IsNullOrWhiteSpace(xmlContent))
            {
                ViewBag.Error = "XML içeriği boş gönderildi.";
                ViewBag.XmlContent = "";
                return View("Index");
            }

            try
            {
                byte[] pdfBytes = ExcelPdfConverter.ConvertToPdf(xmlContent);
                return File(pdfBytes, "application/pdf");
            }
            catch (Exception ex)
            {
                ViewBag.Error = ex.Message;
                ViewBag.XmlContent = xmlContent;
                return View("Index");
            }
        }
    }
}
