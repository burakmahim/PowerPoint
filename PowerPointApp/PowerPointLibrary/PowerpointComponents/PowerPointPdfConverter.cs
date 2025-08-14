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
using PowerPointLibrary.PowerpointComponents;

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

namespace PowerPointLibrary.PowerpointComponents
{
    public static class PowerPointPdfConverter
    {
        public static byte[] ConvertToPdf(string xmlContent)
        {
            byte[] pptxBytes = PowerPointGenerator.CreatePresentationFromXml(xmlContent);
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
