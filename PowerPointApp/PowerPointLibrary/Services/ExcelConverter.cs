using System.Data;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using System.Collections.Generic;
using PowerPointLibrary.Exceptions;
using Syncfusion.Pdf;
using Syncfusion.XlsIO;


#if NET48
using Syncfusion.ExcelToPdfConverter;
#elif NET9_0
using Syncfusion.XlsIORenderer;
#endif

namespace PowerPointLibrary.Services
{
    public static class ExcelConverter
    {
        public static byte[] ConvertToPdf(string xmlContent)
        {
            byte[] excelBytes = ExcelLibrary.CreateExcelFromCustomXml(xmlContent);

            using MemoryStream ms = new MemoryStream(excelBytes);

#if NET48
            using ExcelEngine excelEngine = new ExcelEngine();
            IApplication application = excelEngine.Excel;
            application.DefaultVersion = ExcelVersion.Xlsx;

            IWorkbook workbook = application.Workbooks.Open(ms);

            ExcelToPdfConverter converter = new ExcelToPdfConverter(workbook);

            ExcelToPdfConverterSettings settings = new ExcelToPdfConverterSettings
            {
                LayoutOptions = LayoutOptions.FitSheetOnOnePage
            };

            PdfDocument pdfDocument = converter.Convert(settings);

            using MemoryStream outMs = new MemoryStream();
            pdfDocument.Save(outMs);
            return outMs.ToArray();

#elif NET9_0
    using ExcelEngine excelEngine = new ExcelEngine();
    IApplication application = excelEngine.Excel;
    application.DefaultVersion = ExcelVersion.Xlsx;

    IWorkbook workbook = application.Workbooks.Open(ms);

    XlsIORendererSettings settings = new XlsIORendererSettings
    {
        LayoutOptions = LayoutOptions.FitSheetOnOnePage
    };

    XlsIORenderer renderer = new XlsIORenderer();
    PdfDocument pdfDocument = renderer.ConvertToPDF(workbook, settings);

    using MemoryStream pdfStream = new MemoryStream();
    pdfDocument.Save(pdfStream);

    return pdfStream.ToArray();

#else
    throw new PlatformNotSupportedException("Bu platform desteklenmiyor.");
#endif
        }
    }
}
