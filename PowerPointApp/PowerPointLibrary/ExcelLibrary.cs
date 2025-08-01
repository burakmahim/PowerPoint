using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using System.Collections.Generic;
using PowerPointLibrary.Exceptions;
using Syncfusion.Pdf;
using Syncfusion.XlsIO;
using Syncfusion.Presentation;
using Syncfusion.XlsIO.Parser.Biff_Records;
using PowerPointLibrary.ExcelHelpers;


#if NET48
using Syncfusion.ExcelToPdfConverter;
using Syncfusion.ExcelChartToImageConverter;
#elif NET9_0
using Syncfusion.XlsIORenderer;
#endif

namespace PowerPointLibrary
{
    public static class ExcelLibrary
    {
        public static byte[] CreateExcelFromCustomXml(string xmlContent)
        {

            try
            {
                using ExcelEngine excelEngine = new ExcelEngine();
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;

                IWorkbook workbook = application.Workbooks.Create(0);

                int pageCounter = 1;
                XElement document = XElement.Parse(xmlContent);

                foreach (XElement sheetXml in document.Descendants("sheet"))
                {
                    string? sheetName = sheetXml.Attribute("name")?.Value ?? $"Sayfa{pageCounter}";
                    IWorksheet sheet = workbook.Worksheets.Create(sheetName);
                    sheet.EnableSheetCalculations();

                    IRange startRange = sheet.Range["A1"];

                    int rowCount = 0;
                    int colCount = 0;

                    IEnumerable<XElement> tables = sheetXml.Elements("table");

                    if (tables != null)
                    {
                        foreach (XElement table in tables)
                        {
                            TableBuilder.AddTable(table, sheet, out rowCount, out colCount);
                        }
                    }

                    IEnumerable<XElement> charts = sheetXml.Elements("chart");

                    int currentTopRow = startRange.Row + rowCount + 1;

                    foreach (XElement chart in charts)
                    {
                        int chartWidth = int.TryParse(chart.Attribute("chartWidth")?.Value, out int w) ? w : 10;
                        int chartHeight = int.TryParse(chart.Attribute("chartHeight")?.Value, out int h) ? h : 15;

                        ChartBuilder.AddChart(chart, sheet, startRange, rowCount, colCount, currentTopRow, chartWidth, chartHeight);

                        currentTopRow += chartHeight + 1;
                    }

                    sheet.UsedRange.AutofitColumns();
                    sheet.UsedRange.AutofitRows();
                    sheet.Calculate();

                    pageCounter++;
                }

                using MemoryStream ms = new MemoryStream();
                workbook.SaveAs(ms);
                return ms.ToArray();

            }
            catch (Exception ex)
            {
                throw new ExcelGenerationException("Excel oluşturulurken bir hata meydana geldi.", ex);
            }

        }


        public static byte[] ConvertToPdf(string xmlContent)
        {
            byte[] excelBytes = CreateExcelFromCustomXml(xmlContent);

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
            converter.ChartToImageConverter = new ChartToImageConverter();

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
