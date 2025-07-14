using System;
using System.IO;
using System.Linq;
using System.Net;
using System.Xml.Linq;
using System.Text;
using Syncfusion.Pdf;

#if !NET48
using Syncfusion.Presentation;
using Syncfusion.Drawing;
using Syncfusion.PresentationToPdfConverter;
using Syncfusion.OfficeChart;

#elif !NET9_0
using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;
using Syncfusion.OfficeChart;
#endif


namespace PowerPointLibrary
{
    public static class PowerPointGenerator
    {
        public static byte[] CreatePresentationFromXml(string xmlContent)
        {
            try
            {
                var doc = XDocument.Parse(xmlContent);
                using var ppt = Presentation.Create();

                var settings = doc.Root?.Element("settings");
                if (settings != null)
                {
                    var masterBg = settings.Element("masterBackground");
                    if (masterBg != null)
                    {
                        var master = ppt.Masters[0];
                        master.Background.Fill.FillType = FillType.Solid;
                        master.Background.Fill.SolidFill.Color = ParseColor(masterBg);
                    }
                }

                foreach (var slideEl in doc.Root.Elements("slide"))
                {
                    var slide = ppt.Slides.Add(SlideLayoutType.Blank);

                    var bg = slideEl.Element("background");
                    if (bg != null)
                    {
                        slide.Background.Fill.FillType = FillType.Solid;
                        slide.Background.Fill.SolidFill.Color = ParseColor(bg);
                    }

                    var title = slideEl.Element("title");
                    if (title != null)
                    {
                        AddTextBox(slide, title.Value,
                            (int?)title.Attribute("x") ?? 50,
                            (int?)title.Attribute("y") ?? 50,
                            (int?)title.Attribute("w") ?? 600,
                            (int?)title.Attribute("h") ?? 50,
                            bold: true, fontSize: 24);
                    }

                    var body = slideEl.Element("body");
                    if (body != null)
                    {
                        AddTextBox(slide, body.Value,
                            (int?)body.Attribute("x") ?? 50,
                            (int?)body.Attribute("y") ?? 120,
                            (int?)body.Attribute("w") ?? 600,
                            (int?)body.Attribute("h") ?? 300);
                    }

                    var footer = slideEl.Element("footer");
                    if (footer != null)
                    {
                        AddTextBox(slide, footer.Value,
                            (int?)footer.Attribute("x") ?? 50,
                            (int?)footer.Attribute("y") ?? 680,
                            (int?)footer.Attribute("w") ?? 600,
                            (int?)footer.Attribute("h") ?? 30,
                            alignment: HorizontalAlignmentType.Center, fontSize: 12);
                    }

                    var table = slideEl.Element("table");
                    if (table != null)
                        AddTable(slide, table);

                    //var videoEl = slides[i].Element("video");
                    //if (videoEl != null)
                    //    AddVideoPlaceholderImage(pres.Slides[i], videoEl);


                    var img = slideEl.Element("img");
                    if (img != null)
                        AddImage(slide, img);

                    Action<string, ListType> AddList = (tag, type) =>
                    {
                        var el = slideEl.Element(tag);
                        if (el != null)
                        {
                            var box = slide.AddTextBox(
                                (int?)el.Attribute("x") ?? 50,
                                (int?)el.Attribute("y") ?? 220,
                                (int?)el.Attribute("w") ?? 600,
                                (int?)el.Attribute("h") ?? 100
                            );
                            foreach (var li in el.Elements("li"))
                            {
                                var p = box.TextBody.AddParagraph(li.Value);
                                p.ListFormat.Type = type;
                                if (type == ListType.Numbered)
                                    p.ListFormat.NumberStyle = NumberedListStyle.ArabicPeriod;
                                p.FirstLineIndent = -20;
                                p.LeftIndent = 20;
                            }
                        }
                    };
                    AddList("ul", ListType.Bulleted);
                    AddList("ol", ListType.Numbered);
                }

                using var ms = new MemoryStream();
                ppt.Save(ms);
                return ms.ToArray();
            }
            catch (Exception ex)
            {
                throw new Exception($"Sunum oluşturulamadı: {ex.Message}");
            }
        }




        public static byte[] ConvertToPdf(string xmlContent)
        {
            byte[] pptxBytes = CreatePresentationFromXml(xmlContent);
            using var ms = new MemoryStream(pptxBytes);

            #if !NET48
                                    using (IPresentation presentation = Presentation.Open(ms))
                                    {
                                        PdfDocument pdfDocument = PresentationToPdfConverter.Convert(presentation);

                                        using MemoryStream outMs = new MemoryStream();
                                        pdfDocument.Save(outMs);
                                        return outMs.ToArray();
                                    }
            #elif !NET9_0
                            throw new NotSupportedException(".NET 9.0 altında PDF'e dönüştürme desteklenmiyor.");
            #else
                            throw new PlatformNotSupportedException("Bu platform desteklenmiyor.");
            #endif
        }



        #region Yardımcı Metodlar

        private static ColorObject ParseColor(XElement element)
        {
            var r = byte.TryParse((string)element.Attribute("r"), out var rVal) ? rVal : (byte)255;
            var g = byte.TryParse((string)element.Attribute("g"), out var gVal) ? gVal : (byte)255;
            var b = byte.TryParse((string)element.Attribute("b"), out var bVal) ? bVal : (byte)255;

            return (ColorObject)ColorObject.FromArgb(r, g, b);
        }


        private static void AddTextBox(ISlide slide, string text, int x, int y, int width, int height,
            bool bold = false, int fontSize = 12, HorizontalAlignmentType alignment = HorizontalAlignmentType.Left)
        {
            var textBox = slide.AddTextBox(x, y, width, height);
            var paragraph = textBox.TextBody.AddParagraph(text);
            paragraph.HorizontalAlignment = alignment;
            paragraph.Font.Bold = bold;
            paragraph.Font.FontSize = fontSize;
        }

        private static void AddTable(ISlide slide, XElement tableElement)
        {
            var rows = tableElement.Elements("tr").ToList();
            if (rows.Count == 0) return;

            int rowCount = rows.Count;
            int colCount = rows[0].Elements("td").Count();

            var table = slide.Shapes.AddTable(rowCount, colCount,
                (int?)tableElement.Attribute("x") ?? 50,
                (int?)tableElement.Attribute("y") ?? 400,
                (int?)tableElement.Attribute("w") ?? 600,
                rowCount * 30 + 20);

            for (int r = 0; r < rowCount; r++)
            {
                var cells = rows[r].Elements("td").ToList();
                for (int c = 0; c < colCount; c++)
                {
                    table.Rows[r].Cells[c].TextBody.AddParagraph(c < cells.Count ? cells[c].Value : "");
                }
            }
        }


        private static void AddImage(ISlide slide, XElement imgElement)
        {
            var imagePath = imgElement.Attribute("src")?.Value ?? imgElement.Attribute("url")?.Value;
            if (string.IsNullOrWhiteSpace(imagePath)) return;

            try
            {
                byte[] imageBytes;

                if (imagePath.StartsWith("http", StringComparison.OrdinalIgnoreCase))
                {
                    using var webClient = new WebClient();
                    imageBytes = webClient.DownloadData(imagePath);
                }
                else if (File.Exists(imagePath))
                {
                    imageBytes = File.ReadAllBytes(imagePath);
                }
                else
                {
                    return;
                }

                using var stream = new MemoryStream(imageBytes);
                slide.Pictures.AddPicture(
                    stream,
                    (int?)imgElement.Attribute("x") ?? 50,
                    (int?)imgElement.Attribute("y") ?? 200,
                    (int?)imgElement.Attribute("w") ?? 300,
                    (int?)imgElement.Attribute("h") ?? 200);
            }
            catch
            {
                // Hata durumunda sessizce devam et
            }
        }

        #endregion
    }
}
