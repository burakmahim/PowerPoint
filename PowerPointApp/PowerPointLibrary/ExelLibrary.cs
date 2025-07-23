using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using System.Collections.Generic;
using PowerPointLibrary.Helpers;
using PowerPointLibrary.Services;
using PowerPointLibrary.Exceptions;




#if NET48 || NET9_0
using Syncfusion.XlsIO;
#endif

namespace PowerPointLibrary
{
    public static class ExcelLibrary
    {

        public static byte[] CreateExcelFromCustomXml(string xmlContent)
        {
#if NET48 || NET9_0
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

                    foreach (XElement table in sheetXml.Elements("table"))
                    {
                        currentRow = ExcelTableBuilder.AddTable(table, sheet, currentRow, currentColumn, tableMap);
                        currentRow += 1;
                    }

                    foreach (XElement chartElement in sheetXml.Elements("chart"))
                    {
                        ExcelChartBuilder.AddChart(chartElement, sheet, currentRow, tableMap);
                        currentRow += 22;
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
#else
    throw new PlatformNotSupportedException("Bu platform desteklenmiyor.");
#endif
        }


    }
}