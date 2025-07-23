using Microsoft.AspNetCore.Mvc;
using System.Text;
using PowerPointLibrary;
using PowerPointLibrary.Exceptions;


namespace PowerPointApp.Core.Controllers
{
    [Route("[controller]/[action]")]
    public class PresentationController : Controller
    {
        [HttpGet]
        public IActionResult Index()
        {
            return View(); // Views/Presentation/Index.cshtml
        }

        [HttpPost]
        public IActionResult DownloadPptx([FromForm] string xmlContent)
        {
            if (string.IsNullOrWhiteSpace(xmlContent))
                return BadRequest("XML içeriği boş olamaz.");

            try
            {
                var pptBytes = PowerPointGenerator.CreatePresentationFromXml(xmlContent);
                return File(
                    pptBytes,
                    "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    "Sunum.pptx"
                );
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"PowerPoint oluşturulamadı: {ex.Message}");
            }
        }

        [HttpPost]
        public IActionResult GenerateExcelFromXml([FromForm] string xmlContent)
        {
            if (string.IsNullOrWhiteSpace(xmlContent))
                return BadRequest("XML içeriği boş olamaz.");

            try
            {
                byte[] result = ExcelLibrary.CreateExcelFromCustomXml(xmlContent);
                return File(
                    result,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    "veriler.xlsx"
                );
            }
            catch (ExcelGenerationException ex)
            {
                return StatusCode(500, $"Excel hatası: {ex.Message}");
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Beklenmeyen hata: {ex.Message}");
            }
        }

        [HttpPost]
        public IActionResult Generate([FromForm] string xmlContent, [FromForm] string format)
        {
            if (string.IsNullOrWhiteSpace(xmlContent))
                return BadRequest("XML içeriği boş olamaz.");

            try
            {
                byte[] result;
                string mime, filename;

                if (format?.ToLowerInvariant() == "pdf")
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
            catch (Exception ex)
            {
                return StatusCode(500, $"Dosya oluşturulamadı: {ex.Message}");
            }
        }
    }
}