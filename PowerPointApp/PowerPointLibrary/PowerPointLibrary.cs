using System;
using System.IO;
using System.Linq;
using System.Net;
using System.Xml.Linq;
using System.Text;
using Syncfusion.Pdf;

#if NET48
using Syncfusion.Presentation;
using Syncfusion.Drawing;
using Syncfusion.PresentationToPdfConverter;
using Syncfusion.OfficeChart;

#elif NET9_0
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
                XElement document = XElement.Parse(xmlContent);
                using IPresentation presentation = Presentation.Create();

                XElement? settings = document.Element("settings");
                if (settings != null)
                {
                    string? masterBgColor = settings.Attribute("masterBackgroundColor")?.Value;
                    if (masterBgColor != null)
                    {
                        IMasterSlide slideMaster = presentation.Masters[0];
                        slideMaster.Background.Fill.FillType = FillType.Solid;
                        slideMaster.Background.Fill.SolidFill.Color = ParseColor(masterBgColor);
                    }
                }

                foreach (XElement slideElement in document.Elements("slide"))
                {
                    SlideLayoutType slideLayoutType = Enum.TryParse(slideElement.Attribute("layout")?.Value, true, out SlideLayoutType lt) ? lt : SlideLayoutType.TitleAndContent;

                    ISlide slide = presentation.Slides.Add(slideLayoutType);

                    string? title        = slideElement.Element("title")?.Value;
                    string? subtitle     = slideElement.Element("subtitle")?.Value;
                    string? content      = slideElement.Element("content")?.Value;
                    string? leftcontent  = slideElement.Element("leftcontent")?.Value;
                    string? rightcontent = slideElement.Element("rightcontent")?.Value;

                    foreach (IShape shape in slide.Shapes)
                    {
                        switch (slideLayoutType)
                        {
                            case SlideLayoutType.TitleAndContent:
                                if (shape.PlaceholderFormat.Type == PlaceholderType.Title)
                                {
                                    if (!string.IsNullOrEmpty(title))
                                        shape.TextBody.AddParagraph(title);
                                }
                                else if (shape.PlaceholderFormat.Type == PlaceholderType.Body)
                                {
                                    if (!string.IsNullOrEmpty(content))
                                        shape.TextBody.AddParagraph(content);
                                }
                                break;

                            case SlideLayoutType.TitleOnly:
                                if (shape.PlaceholderFormat.Type == PlaceholderType.Title)
                                {
                                    if (!string.IsNullOrEmpty(title))
                                        shape.TextBody.AddParagraph(title);
                                }
                                break;

                            case SlideLayoutType.TwoContent:
                                if (shape.PlaceholderFormat.Type == PlaceholderType.Title)
                                {
                                    if (!string.IsNullOrEmpty(title))
                                        shape.TextBody.AddParagraph(title);
                                }
                                else if (shape.PlaceholderFormat.Type == PlaceholderType.Body)
                                {
                                    if (!string.IsNullOrEmpty(leftcontent))
                                        shape.TextBody.AddParagraph(leftcontent);
                                    if (!string.IsNullOrEmpty(rightcontent))
                                        shape.TextBody.AddParagraph(rightcontent);
                                }
                                break;

                            case SlideLayoutType.SectionHeader:
                                if (shape.PlaceholderFormat.Type == PlaceholderType.Title)
                                {
                                    if (!string.IsNullOrEmpty(title))
                                        shape.TextBody.AddParagraph(title);
                                }
                                else if (shape.PlaceholderFormat.Type == PlaceholderType.Subtitle)
                                {
                                    if (!string.IsNullOrEmpty(subtitle))
                                        shape.TextBody.AddParagraph(subtitle);
                                }
                                break;

                            case SlideLayoutType.Blank:
                                break;

                            default:
                                if (shape.PlaceholderFormat.Type == PlaceholderType.Title)
                                {
                                    if (!string.IsNullOrEmpty(title))
                                        shape.TextBody.AddParagraph(title);
                                }
                                else if (shape.PlaceholderFormat.Type == PlaceholderType.Body)
                                {
                                    if (!string.IsNullOrEmpty(content))
                                        shape.TextBody.AddParagraph(content);
                                }
                                break;
                        }
                    }

                    string? slideBackgroundColor = slideElement.Attribute("backgroundColor")?.Value;
                    if (slideBackgroundColor != null)
                    {
                        slide.Background.Fill.FillType = FillType.Solid;
                        slide.Background.Fill.SolidFill.Color = ParseColor(slideBackgroundColor);
                    }

                    XElement? imageElement = slideElement.Element("image");
                    if(imageElement != null)
                    {
                        double x = (double.TryParse(imageElement.Attribute("x")?.Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var dx) ? dx : 1) * 28.3465;
                        double y = (double.TryParse(imageElement.Attribute("y")?.Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var dy) ? dy : 1) * 28.3465;
                        double cx = (double.TryParse(imageElement.Attribute("w")?.Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var dw) ? dw : 5) * 28.3465;
                        double cy = (double.TryParse(imageElement.Attribute("h")?.Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var dh) ? dh : 5) * 28.3465;

                        AddImage(slide, imageElement, x, y, cx, cy);
                    }

                    XElement? tableElement = slideElement.Element("table");
                    if (tableElement != null)
                    {
                        AddTable(slide, tableElement);
                    }

                    XElement? textboxElement = slideElement.Element("textbox");
                    if(textboxElement != null)
                    {

                        string? text = textboxElement.Value;

                        double x  = (double.TryParse(textboxElement.Attribute("x")?.Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var dx) ? dx : 1) * 28.3465;
                        double y  = (double.TryParse(textboxElement.Attribute("y")?.Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var dy) ? dy : 1) * 28.3465;
                        double cx = (double.TryParse(textboxElement.Attribute("w")?.Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var dw) ? dw : 5) * 28.3465;
                        double cy = (double.TryParse(textboxElement.Attribute("h")?.Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var dh) ? dh : 5) * 28.3465;

                        AutoShapeType shapeType = Enum.TryParse<AutoShapeType>(textboxElement.Attribute("shapeType")?.Value ?? "Rectangle", true, out var st) ? st : AutoShapeType.Rectangle;

                        bool bold    = bool.TryParse(textboxElement.Attribute("bold")?.Value, out var b) && b;
                        bool italic  = bool.TryParse(textboxElement.Attribute("italic")?.Value, out var i) && i;

                        string? textColor       = textboxElement.Attribute("textColor")?.Value ?? "#000000";
                        string? backgroundColor = textboxElement.Attribute("backgroundColor")?.Value;
                        int fontSize            = int.TryParse(textboxElement.Attribute("fontSize")?.Value, out var fs) ? fs : 12;
                       
                        HorizontalAlignmentType alignment = Enum.TryParse(textboxElement.Attribute("alignment")?.Value, true, out HorizontalAlignmentType align) ? align : HorizontalAlignmentType.Left;

                        AddShape(slide, text, shapeType, x, y, cx, cy, bold, italic, textColor, backgroundColor, fontSize, alignment);
                    }

                    //XElement? chartElement = slideElement.Element("chart");
                    //if(chartElement!= null)
                    //{
                    //    string chartTypeStr = chartElement.Attribute("type").Value;
                    //    Chart chartType = Enum.TryParse(chartTypeStr, true, out ChartType ct) ? ct : ChartType.Pie;
                    //}

                    Action<string, ListType> AddList = (tag, type) =>
                    {
                        XElement el = slideElement.Element(tag);
                        if (el != null)
                        {
                            IShape box = slide.AddTextBox(
                                (int?)el.Attribute("x") ?? 50,
                                (int?)el.Attribute("y") ?? 220,
                                (int?)el.Attribute("w") ?? 600,
                                (int?)el.Attribute("h") ?? 100
                            );
                            foreach (var li in el.Elements("li"))
                            {
                                IParagraph p = box.TextBody.AddParagraph(li.Value);
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

                using MemoryStream ms = new MemoryStream();
                presentation.Save(ms);
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

            #if NET48
                                    using (IPresentation presentation = Presentation.Open(ms))
                                    {
                                        PdfDocument pdfDocument = PresentationToPdfConverter.Convert(presentation);

                                        using MemoryStream outMs = new MemoryStream();
                                        pdfDocument.Save(outMs);
                                        return outMs.ToArray();
                                    }
            #elif NET9_0
                            throw new NotSupportedException(".NET 9.0 altında PDF'e dönüştürme desteklenmiyor.");
            #else
                            throw new PlatformNotSupportedException("Bu platform desteklenmiyor.");
            #endif
        }



        private static ColorObject ParseColor(string hexColor)
        {
            if (hexColor.StartsWith("#"))
                hexColor = hexColor.Substring(1);

            byte r = Convert.ToByte(hexColor.Substring(0, 2), 16);
            byte g = Convert.ToByte(hexColor.Substring(2, 2), 16);
            byte b = Convert.ToByte(hexColor.Substring(4, 2), 16);

            return (ColorObject)ColorObject.FromArgb(r, g, b);
        }
        private static void AddShape(ISlide slide, string text, AutoShapeType shapeType, double x, double y, double width, double height, bool bold = false, bool italic = false, string textColor = "#000000", string? backgroundColor = null, int fontSize = 12, HorizontalAlignmentType alignment = HorizontalAlignmentType.Left)
        {
            IShape shape = slide.Shapes.AddShape( shapeType, x, y, width, height);
            shape.Fill.FillType = FillType.None;

            IParagraph paragraph = shape.TextBody.AddParagraph(text);
            paragraph.HorizontalAlignment = alignment;
            paragraph.Font.Bold = bold;
            paragraph.Font.Italic = italic;
            paragraph.Font.FontSize = fontSize;
            paragraph.Font.Color = ParseColor(textColor);


            if (backgroundColor != null)
            {
                shape.Fill.FillType = FillType.Solid;
                shape.Fill.SolidFill.Color = ParseColor(backgroundColor);
            }
        }
        private static void AddTable(ISlide slide, XElement tableElement)
        {
            List<XElement> rows = tableElement.Elements("tr").ToList();
            if (rows.Count == 0) return;

            int rowCount = rows.Count;
            int colCount = rows[0].Elements("td").Count();

            ITable table = slide.Shapes.AddTable(rowCount, colCount,
                (int?)tableElement.Attribute("x") ?? 50,
                (int?)tableElement.Attribute("y") ?? 400,
                (int?)tableElement.Attribute("w") ?? 600,
                rowCount * 30 + 20);

            for (int r = 0; r < rowCount; r++)
            {
                List<XElement> cells = rows[r].Elements("td").ToList();
                for (int c = 0; c < colCount; c++)
                {
                    table.Rows[r].Cells[c].TextBody.AddParagraph(c < cells.Count ? cells[c].Value : "");
                }
            }
        }


        private static void AddImage(ISlide slide, XElement imgElement, double x, double y, double cx, double cy)
        {
            string? imagePath = imgElement.Attribute("path")?.Value;
            
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
                else if(imagePath.StartsWith("data:"))
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
