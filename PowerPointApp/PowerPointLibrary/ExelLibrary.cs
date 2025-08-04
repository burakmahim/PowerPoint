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
using PowerPointLibrary.ExcelServices;

//using Syncfusion.PresentationToPdfConverter;



#if NET48
using Syncfusion.ExcelToPdfConverter;
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
                application.DefaultVersion = ExcelVersion.Excel2016;

                IWorkbook workbook = application.Workbooks.Create(0);
                int pageCounter = 1;
                var document = XElement.Parse(xmlContent);

                foreach (XElement sheetXml in document.Descendants("sheet"))
                {
                    int currentRow = 1;
                    int currentColumn = 1;

                    string sheetName = sheetXml.Attribute("name")?.Value ?? $"Sayfa{pageCounter}";
                    IWorksheet sheet = workbook.Worksheets.Create(sheetName);

                    var tableMap = new Dictionary<string, (int StartRow, int StartCol, int RowCount, int ColCount)>();

                    // Sırayla tüm alt elemanları işle
                    foreach (XElement element in sheetXml.Elements())
                    {
                        switch (element.Name.LocalName)
                        {
                            case "table":
                                currentRow = ExcelTableBuilder.AddTable(element, sheet, currentRow, currentColumn, tableMap);
                                currentRow += 1; // tablo sonrası boşluk
                                break;

                            case "chart":
                                ExcelChartBuilder.AddChart(element, sheet, currentRow, tableMap);
                                currentRow += 22; // grafik yüksekliği kadar boşluk
                                break;

                            // Eğer başka özel elemanlar varsa buraya eklenebilir
                            default:
                                // Bilinmeyen bir eleman varsa geç ve satır atla
                                currentRow += 1;
                                break;
                        }
                    }

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

        public static (IWorkbook, IWorksheet, Dictionary<string, (int, int, int, int)>) CreateWorkbookFromXml(string xmlContent)
        {
            ExcelEngine excelEngine = new ExcelEngine();
            IApplication application = excelEngine.Excel;
            application.DefaultVersion = ExcelVersion.Xlsx;

            IWorkbook workbook = application.Workbooks.Create(1);
            IWorksheet sheet = workbook.Worksheets[0];

            XElement document = XElement.Parse(xmlContent);
            Dictionary<string, (int, int, int, int)> tableMap = new();

            int currentRow = 1;
            int currentCol = 1;

            foreach (XElement element in document.Elements("sheet").Elements())
            {
                if (element.Name == "table")
                {
                    currentRow = ExcelTableBuilder.AddTable(element, sheet, currentRow, currentCol, tableMap);
                    currentRow += 1;
                }
                else if (element.Name == "chart")
                {
                    ExcelChartBuilder.AddChart(element, sheet, currentRow, tableMap);
                    currentRow += 22;
                }
            }

            return (workbook, sheet, tableMap);
        }



    }
}
