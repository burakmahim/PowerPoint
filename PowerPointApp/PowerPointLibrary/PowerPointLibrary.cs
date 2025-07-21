using System;
using System.IO;
using System.Linq;
using System.Net;
using System.Xml.Linq;
using System.Text;
using Syncfusion.Pdf;
using System.Drawing;
using System.Security.Policy;
using System.Collections.Generic;



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

                XElement? footerElement = document.Element("footer");


                foreach (XElement slideElement in document.Elements("slide"))
                {
                    SlideLayoutType slideLayoutType = Enum.TryParse(slideElement.Attribute("layout")?.Value, true, out SlideLayoutType lt) ? lt : SlideLayoutType.TitleAndContent;

                    ISlide slide = presentation.Slides.Add(slideLayoutType);

                    string? title        = slideElement.Element("title")?.Value;
                    string? subtitle     = slideElement.Element("subtitle")?.Value;
                    string? content      = slideElement.Element("content")?.Value;
                    string? leftcontent  = slideElement.Element("leftcontent")?.Value;
                    string? rightcontent = slideElement.Element("rightcontent")?.Value;

                    int bodyCounter = 0;

                    foreach (IShape shape in slide.Shapes)
                    {
                        switch (slideLayoutType)
                        {
                            case SlideLayoutType.TitleAndContent:
                                if (shape.PlaceholderFormat.Type == PlaceholderType.Title && !string.IsNullOrEmpty(title))
                                    shape.TextBody.AddParagraph(title);
                                else if (shape.PlaceholderFormat.Type == PlaceholderType.Body && !string.IsNullOrEmpty(content))
                                    shape.TextBody.AddParagraph(content);
                                break;

                            case SlideLayoutType.TitleOnly:
                                if (shape.PlaceholderFormat.Type == PlaceholderType.Title && !string.IsNullOrEmpty(title))
                                    shape.TextBody.AddParagraph(title);
                                break;

                            case SlideLayoutType.TwoContent:
                                if (shape.PlaceholderFormat.Type == PlaceholderType.Title && !string.IsNullOrEmpty(title))
                                {
                                    shape.TextBody.AddParagraph(title);
                                }
                                else if (shape.PlaceholderFormat.Type == PlaceholderType.Body)
                                {
                                    if (slideLayoutType == SlideLayoutType.TwoContent)
                                    {
                                        if (bodyCounter == 0 && !string.IsNullOrEmpty(leftcontent))
                                            shape.TextBody.AddParagraph(leftcontent);
                                        else if (bodyCounter == 1 && !string.IsNullOrEmpty(rightcontent))
                                            shape.TextBody.AddParagraph(rightcontent);

                                        bodyCounter++;
                                    }
                                    else if (!string.IsNullOrEmpty(content))
                                    {
                                        shape.TextBody.AddParagraph(content);
                                    }
                                }

                                break;

                            case SlideLayoutType.SectionHeader:
                                if (shape.PlaceholderFormat.Type == PlaceholderType.Title && !string.IsNullOrEmpty(title))
                                {
                                    shape.TextBody.AddParagraph(title);
                                }
                                else if ((shape.PlaceholderFormat.Type == PlaceholderType.Subtitle ||
                                          shape.PlaceholderFormat.Type == PlaceholderType.Body) &&
                                          !string.IsNullOrEmpty(subtitle))
                                {
                                    shape.TextBody.AddParagraph(subtitle);
                                }
                                break;

                            case SlideLayoutType.Blank:
                                break;

                            default:
                                if (shape.PlaceholderFormat.Type == PlaceholderType.Title && !string.IsNullOrEmpty(title))
                                    shape.TextBody.AddParagraph(title);
                                else if (shape.PlaceholderFormat.Type == PlaceholderType.Body && !string.IsNullOrEmpty(content))
                                    shape.TextBody.AddParagraph(content);
                                break;
                        }
                    }

                    if (footerElement != null)
                    {
                        SetHeaderFooter(document, slide);
                    }

                    string? slideBackgroundColor = slideElement.Attribute("backgroundColor")?.Value;
                    if (slideBackgroundColor != null)
                    {
                        slide.Background.Fill.FillType = FillType.Solid;
                        slide.Background.Fill.SolidFill.Color = ParseColor(slideBackgroundColor);
                    }

                    IEnumerable <XElement> imageElements = slideElement.Elements("image");
                    if(imageElements != null)
                    {
                        foreach(XElement imgElement in imageElements)
                        {
                            AddImage(slide, imgElement);
                        }
                    }

                    IEnumerable<XElement> tableElements = slideElement.Elements("table");
                    if (tableElements != null)
                    {
                        foreach (XElement tableElement in tableElements)
                        {
                            AddTable(slide, tableElement);
                        }
                    }

                    IEnumerable<XElement> shapeElements = slideElement.Elements("shape");
                    if(shapeElements != null)
                    {
                        foreach(XElement shapeElement in shapeElements)
                        {
                            string? text = shapeElement.Value;

                            AutoShapeType shapeType = Enum.TryParse<AutoShapeType>(shapeElement.Attribute("shapeType")?.Value ?? "Rectangle", true, out AutoShapeType st) ? st : AutoShapeType.Rectangle;

                            bool bold = bool.TryParse(shapeElement.Attribute("bold")?.Value, out bool b) && b;
                            bool italic = bool.TryParse(shapeElement.Attribute("italic")?.Value, out bool i) && i;

                            string? textColor = shapeElement.Attribute("textColor")?.Value ?? "#000000";
                            string? backgroundColor = shapeElement.Attribute("backgroundColor")?.Value;
                            int fontSize = int.TryParse(shapeElement.Attribute("fontSize")?.Value, out int fs) ? fs : 12;

                            HorizontalAlignmentType alignment = Enum.TryParse(shapeElement.Attribute("alignment")?.Value, true, out HorizontalAlignmentType align) ? align : HorizontalAlignmentType.Left;

                            AddShape(shapeElement, slide, text, shapeType, bold, italic, textColor, backgroundColor, fontSize, alignment);
                        }
                    }

                    IEnumerable<XElement> listElements = slideElement.Elements("list");
                    if(listElements != null)
                    {
                        foreach(XElement listElement in listElements)
                        {
                            AddList(listElement, slide);
                        }
                    }
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
            using MemoryStream ms = new MemoryStream(pptxBytes);

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
        private static void AddShape(XElement shapeElement, ISlide slide, string text, AutoShapeType shapeType, bool bold = false, bool italic = false, string textColor = "#000000", string? backgroundColor = null, int fontSize = 12, HorizontalAlignmentType alignment = HorizontalAlignmentType.Left)
        {
            string? fontFamily = shapeElement.Attribute("fontFamily")?.Value;

            double x = (double.TryParse(shapeElement.Attribute("x")?.Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double dx) ? dx : 1) * 28.3465;
            double y = (double.TryParse(shapeElement.Attribute("y")?.Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double dy) ? dy : 1) * 28.3465;
            double cx = (double.TryParse(shapeElement.Attribute("w")?.Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double dw) ? dw : 5) * 28.3465;
            double cy = (double.TryParse(shapeElement.Attribute("h")?.Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double dh) ? dh : 5) * 28.3465;

            IShape shape = slide.Shapes.AddShape( shapeType, x, y, cx, cy);
            shape.Fill.FillType = FillType.None;

            IParagraph paragraph = shape.TextBody.AddParagraph(text);
            paragraph.HorizontalAlignment = alignment;
            paragraph.Font.Bold = bold;
            paragraph.Font.Italic = italic;
            paragraph.Font.FontSize = fontSize;
            paragraph.Font.FontName = fontFamily;
            paragraph.HorizontalAlignment = alignment;
            paragraph.Font.Color = ParseColor(textColor);

            //shape.LineFormat.Fill.FillType = FillType.None;

            if (backgroundColor != null)
            {
                shape.Fill.FillType = FillType.Solid;
                shape.Fill.SolidFill.Color = ParseColor(backgroundColor);
            }

           //eererereerererer




        }
        private static void AddText()
        {

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
        private static void AddList(XElement listElement, ISlide slide)
        {

            double x = (double.TryParse(listElement.Attribute("x")?.Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double dx) ? dx : 1) * 28.3465;
            double y = (double.TryParse(listElement.Attribute("y")?.Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double dy) ? dy : 1) * 28.3465;
            double cx = (double.TryParse(listElement.Attribute("w")?.Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double dw) ? dw : 5) * 28.3465;
            double cy = (double.TryParse(listElement.Attribute("h")?.Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double dh) ? dh : 5) * 28.3465;

            bool bold = bool.TryParse(listElement.Attribute("bold")?.Value, out bool b) && b;
            bool italic = bool.TryParse(listElement.Attribute("italic")?.Value, out bool i) && i;

            string listFontFamily = listElement.Attribute("font")?.Value ?? "Calibri";
            string listTextColor = listElement.Attribute("color")?.Value ?? "#000000";
            string listFontSize = listElement.Attribute("size")?.Value ?? "14";
            string listType = listElement.Attribute("type")?.Value ?? "bulleted";

            ListType listTypeParsed = Enum.TryParse<ListType>(listType, true, out var result) ? result : ListType.Bulleted;

            int listFontSizeParsed = int.TryParse(listFontSize, out int s) ? s : 14;

            IShape listBox = slide.AddTextBox(x, y, cx, cy);

            foreach (XElement item in listElement.Elements("item"))
            {
                IParagraph paragraph = listBox.TextBody.AddParagraph(item.Value);
                paragraph.ListFormat.Type = listTypeParsed;

                string itemFontFamily = item.Attribute("fontFamily")?.Value ?? listFontFamily;
                string itemTextColor = item.Attribute("color")?.Value ?? listTextColor;
                string itemFontSize = item.Attribute("fontSize")?.Value ?? listFontSize;

                int itemFontSizeParsed = int.Parse(itemFontSize);

                paragraph.Font.FontName = itemFontFamily;
                paragraph.Font.FontSize = itemFontSizeParsed;
                paragraph.Font.Color = ParseColor(itemTextColor);
            }

        }
        private static void SetHeaderFooter(XElement document, ISlide slide)
        {
            XElement? footerElement = document.Element("footer");
            if (footerElement == null) return;

            bool enableFooter = bool.TryParse(document.Element("settings")?.Attribute("footer")?.Value, out bool result) && result;
            string? footerText = footerElement.Value;

            slide.HeadersFooters.Footer.Visible = enableFooter;
            slide.HeadersFooters.Footer.Text = footerText;

            bool enableSlideNumber = bool.TryParse(document.Element("settings")?.Attribute("slideNumber")?.Value, out bool resultNumber) && resultNumber;

            slide.HeadersFooters.SlideNumber.Visible = enableSlideNumber ;

            double x = (double.TryParse(footerElement.Attribute("x")?.Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double dx) ? dx : 1) * 28.3465;
            double y = (double.TryParse(footerElement.Attribute("y")?.Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double dy) ? dy : 12.5) * 28.3465;
            double cx = (double.TryParse(footerElement.Attribute("w")?.Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double dw) ? dw : 23) * 28.3465;
            double cy = (double.TryParse(footerElement.Attribute("h")?.Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double dh) ? dh : 1.2) * 28.3465;

        }
        private static void AddImage(ISlide slide, XElement imgElement)
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

        //private static void AddChart(ISlide slide, XElement chartElement)
        //{
        //    string? chartTypeStr = chartElement.Attribute("type")?.Value;
        //    if (string.IsNullOrWhiteSpace(chartTypeStr))
        //    {
        //        throw new Exception("Chart için 'type' niteliği zorunludur ve boş olamaz.");
        //    }

        //    double x = (double.TryParse(chartElement.Attribute("x")?.Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double dx) ? dx : 1) * 28.3465;
        //    double y = (double.TryParse(chartElement.Attribute("y")?.Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double dy) ? dy : 1) * 28.3465;
        //    double cx = (double.TryParse(chartElement.Attribute("w")?.Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double dw) ? dw : 15) * 28.3465;
        //    double cy = (double.TryParse(chartElement.Attribute("h")?.Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double dh) ? dh : 15) * 28.3465;

        //    OfficeChartType chartType;

        //    if (Enum.TryParse<OfficeChartType>(chartTypeStr, true, out var chartTypeParsed))
        //    {
        //        chartType = chartTypeParsed;

        //        if (chartType == OfficeChartType.Pie)
        //        {

        //            IPresentationChart chart = slide.Charts.AddChart(x, y, cx, cy);
        //            chart.ChartType = OfficeChartType.Pie;

        //            chart.ChartTitle = chartElement.Attribute("title")?.Value;

        //        }
        //        else
        //        {

        //        }

        //    }
        //}


    }
}


//Action<string, ListType> AddList = (tag, type) =>
//{
//    XElement? el = slideElement.Element(tag);
//    if (el != null)
//    {
//        IShape box = slide.AddTextBox(
//            (int?)el.Attribute("x") ?? 50,
//            (int?)el.Attribute("y") ?? 220,
//            (int?)el.Attribute("w") ?? 600,
//            (int?)el.Attribute("h") ?? 100
//        );
//        foreach (XElement li in el.Elements("li"))
//        {
//            IParagraph p = box.TextBody.AddParagraph(li.Value);
//            p.ListFormat.Type = type;
//            if (type == ListType.Numbered)
//                p.ListFormat.NumberStyle = NumberedListStyle.ArabicPeriod;
//            p.FirstLineIndent = -20;
//            p.LeftIndent = 20;
//        }
//    }
//};
//AddList("ul", ListType.Bulleted);
//AddList("ol", ListType.Numbered);