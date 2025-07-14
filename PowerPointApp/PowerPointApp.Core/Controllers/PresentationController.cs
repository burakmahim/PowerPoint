using Microsoft.AspNetCore.Mvc;
using System.Text;
using PowerPointLibrary;

namespace PowerPointApp.Core.Controllers
{
    public class PresentationController : Controller
    {
        [HttpGet]
        public IActionResult Index()
        {
            return View(); // Views/Presentation/Index.cshtml
        }


        [HttpPost]
        public IActionResult DownloadPptx(string xmlContent)
        {
            var pptBytes = PowerPointGenerator.CreatePresentationFromXml(xmlContent);
            return File(pptBytes,
                "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                "Sunum.pptx");
        }


        [HttpPost]
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
