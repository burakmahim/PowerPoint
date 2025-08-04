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
using PowerPointLibrary.PowerPointHelpers;

#if NET48
using Syncfusion.Presentation;
using Syncfusion.OfficeChartToImageConverter;
using Syncfusion.Drawing;
using Syncfusion.PresentationToPdfConverter;
using Syncfusion.OfficeChart;
using Syncfusion.ExcelChartToImageConverter;

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
                        slideMaster.Background.Fill.SolidFill.Color = ColorHelper.ParseColor(masterBgColor);
                    }
                }

                XElement? footerElement = document.Element("footer");


                foreach (XElement slideElement in document.Elements("slide"))
                {
                    SlideLayoutType slideLayoutType = Enum.TryParse(slideElement.Attribute("layout")?.Value, true, out SlideLayoutType lt) ? lt : SlideLayoutType.TitleAndContent;

                    ISlide slide = presentation.Slides.Add(slideLayoutType);

                    LayoutHelper.SetLayoutContent(slide, slideElement, slideLayoutType);


                    if (footerElement != null)
                    {
                        HeaderFooterHelper.SetHeaderFooter(document, slide);
                    }

                    string? slideBackgroundColor = slideElement.Attribute("backgroundColor")?.Value;
                    if (slideBackgroundColor != null)
                    {
                        slide.Background.Fill.FillType = FillType.Solid;
                        slide.Background.Fill.SolidFill.Color = ColorHelper.ParseColor(slideBackgroundColor);
                    }

                    IEnumerable<XElement> chartElements = slideElement.Elements("chart");
                    if (chartElements != null)
                    {
                        foreach (XElement chartElement in chartElements)
                        {
                            ChartHelper.AddChart(slide, chartElement);
                        }
                    }

                    IEnumerable<XElement> imageElements = slideElement.Elements("image");
                    if (imageElements != null)
                    {
                        foreach (XElement imgElement in imageElements)
                        {
                            ImageHelper.AddImage(slide, imgElement);
                        }
                    }

                    IEnumerable<XElement> tableElements = slideElement.Elements("table");
                    if (tableElements != null)
                    {
                        foreach (XElement tableElement in tableElements)
                        {
                            TableHelper.AddTable(slide, tableElement);
                        }
                    }

                    IEnumerable<XElement> shapeElements = slideElement.Elements("shape");
                    if (shapeElements != null)
                    {
                        foreach (XElement shapeElement in shapeElements)
                        {


                            ShapeHelper.AddShape(shapeElement, slide);
                        }
                    }

                    IEnumerable<XElement> listElements = slideElement.Elements("list");
                    if (listElements != null)
                    {
                        foreach (XElement listElement in listElements)
                        {
                            ListHelper.AddList(listElement, slide);
                        }
                    }

                    IEnumerable<XElement> textboxElements = slideElement.Elements("textbox");
                    if (textboxElements != null)
                    {
                        foreach (XElement textboxElement in textboxElements)
                        {
                            TextBoxHelper.AddTextBox(textboxElement, slide);
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
                Syncfusion.OfficeChartToImageConverter.ChartToImageConverter chartToImageConverter = new Syncfusion.OfficeChartToImageConverter.ChartToImageConverter();

                presentation.ChartToImageConverter = chartToImageConverter;

                PresentationToPdfConverterSettings settings = new PresentationToPdfConverterSettings();
                settings.ShowHiddenSlides = false;

                using (PdfDocument pdfDocument = PresentationToPdfConverter.Convert(presentation, settings))
                {
                    using MemoryStream outMs = new MemoryStream();
                    pdfDocument.Save(outMs);
                    return outMs.ToArray();
                }
            }
#elif NET9_0
           using (IPresentation presentation = Presentation.Open(ms))
           {
               PdfDocument pdfDocument = PresentationToPdfConverter.Convert(presentation);

               using MemoryStream outMs = new MemoryStream();
               pdfDocument.Save(outMs);
               return outMs.ToArray();
           }
#else
           throw new PlatformNotSupportedException("Bu platform desteklenmiyor.");
#endif
        }


    }
}