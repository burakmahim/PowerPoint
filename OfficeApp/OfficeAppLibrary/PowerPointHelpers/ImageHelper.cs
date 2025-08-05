using Syncfusion.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace OfficeAppLibrary.PowerPointHelpers
{
    public static class ImageHelper
    {
        public static void AddImage(ISlide slide, XElement imgElement)
        {
            string? imagePath = imgElement.Attribute("path")?.Value;

            double x = (double.TryParse(imgElement.Attribute("x")?.Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double dx) ? dx : 1) * 28.3465;
            double y = (double.TryParse(imgElement.Attribute("y")?.Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double dy) ? dy : 1) * 28.3465;
            double cx = (double.TryParse(imgElement.Attribute("w")?.Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double dw) ? dw : 5) * 28.3465;
            double cy = (double.TryParse(imgElement.Attribute("h")?.Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double dh) ? dh : 5) * 28.3465;

            if (string.IsNullOrWhiteSpace(imagePath)) return;

            try
            {
                byte[] imageBytes;

                if (imagePath.StartsWith("http", StringComparison.OrdinalIgnoreCase))
                {
                    using WebClient webClient = new WebClient();
                    imageBytes = webClient.DownloadData(imagePath);
                }
                else if (File.Exists(imagePath))
                {
                    imageBytes = File.ReadAllBytes(imagePath);
                }
                else if (imagePath.StartsWith("data:"))
                {
                    string base64data = imagePath.Substring(imagePath.IndexOf(",") + 1);
                    imageBytes = Convert.FromBase64String(base64data);
                }
                else
                {
                    return;
                }
                using MemoryStream stream = new MemoryStream(imageBytes);
                slide.Pictures.AddPicture(stream, x, y, cx, cy);
            }
            catch
            {
                // Hata durumunda sessizce devam et
            }
        }
    }
}
