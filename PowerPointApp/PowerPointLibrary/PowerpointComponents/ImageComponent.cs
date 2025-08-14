using Syncfusion.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace PowerPointLibrary.PowerpointComponents
{
    public static class ImageComponent
    {
        public static void AddImage(ISlide slide, XElement imgElement)
        {
            string? imagePath = imgElement.Attribute("path")?.Value;

            (double x, double y, double cx, double cy) = CoordinatesParser.CoordinateParser(imgElement, 1, 1, 5, 5);

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

            }
        }
    }
}
