using Syncfusion.XlsIO;
using System;
using System.Collections.Generic;
using System.Data;
using System.Xml.Linq;


namespace PowerPointLibrary.ExcelServices
{
    public static class ExcelParserService
    {
        public static Dictionary<string, DataTable> GetTablesFromXml(string xmlContent)
        {
            var (workbook, sheet, tableMap) = ExcelLibrary.CreateWorkbookFromXml(xmlContent);

            sheet.EnableSheetCalculations();
            sheet.Calculate();

            return ExcelTableParser.ParseAllTables(sheet, tableMap);
        }


        public static IWorkbook CreateChartOnlyWorkbookFromXml(string xmlContent)
        {
            ExcelEngine excelEngine = new ExcelEngine();
            IApplication application = excelEngine.Excel;
            application.DefaultVersion = ExcelVersion.Xlsx;

            IWorkbook workbook = application.Workbooks.Create(1);
            IWorksheet sheet = workbook.Worksheets[0];

            XElement document = XElement.Parse(xmlContent);
            Dictionary<string, (int, int, int, int)> fakeTableMap = new();

            int currentRow = 1;
            foreach (XElement element in document.Elements("sheet").Elements("chart"))
            {
                ExcelChartBuilder.AddChart(element, sheet, currentRow, fakeTableMap);
                currentRow += 22;
            }

            return workbook;
        }

    }
}
