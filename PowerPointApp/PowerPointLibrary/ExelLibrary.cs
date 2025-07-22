using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using System.Collections.Generic;

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

                IEnumerable<XElement> tables = sheetXml.Elements("table");
                if (tables != null)
                {
                    foreach (XElement table in tables)
                    {
                        AddTable(table, sheet, currentRow, currentColumn);
                    }
                }

                XElement chartElement = sheetXml.Element("chart");
                if (chartElement != null)
                {
                    AddChart(chartElement, sheet, sheetXml);
                }

                pageCounter++;
            }

            using MemoryStream ms = new MemoryStream();
            workbook.SaveAs(ms);
            return ms.ToArray();
#else
            throw new PlatformNotSupportedException("Bu platform desteklenmiyor.");
#endif
        }

        private static void AddTable(XElement table, IWorksheet sheet, int startRow, int startCol)
        {
            string? startCell = table.Attribute("startCell")?.Value;
            if (!string.IsNullOrWhiteSpace(startCell))
            {
                IRange range = sheet.Range[startCell];
                startRow = range.Row;
                startCol = range.Column;
            }

            List<XElement> rows = table.Elements("row").ToList();
            for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++)
            {
                List<XElement> cells = rows[rowIndex].Elements("cell").ToList();
                for (int colIndex = 0; colIndex < cells.Count; colIndex++)
                {
                    IRange cell = sheet[startRow + rowIndex, startCol + colIndex];
                    string value = cells[colIndex].Value;

                    if (double.TryParse(value, out double numericValue))
                        cell.Number = numericValue;
                    else
                        cell.Text = value;

                    if (cells[colIndex].Attribute("bold")?.Value == "true")
                        cell.CellStyle.Font.Bold = true;
                }
            }

            sheet.UsedRange.AutofitColumns();
        }

        private static void AddChart(XElement chartElement, IWorksheet sheet, XElement sheetXml)
        {
            var chart = sheet.Charts.Add();
            chart.ChartType = ParseChartType(chartElement.Attribute("type")?.Value ?? "Column");
            chart.ChartTitle = chartElement.Attribute("title")?.Value ?? "";
            chart.PrimaryCategoryAxis.Title = chartElement.Attribute("xAxis")?.Value ?? "";
            chart.PrimaryValueAxis.Title = chartElement.Attribute("yAxis")?.Value ?? "";

            int chartStartRow = 20;
            int chartStartCol = 1;

            var dataRangeAttr = chartElement.Attribute("dataRange")?.Value;
            if (!string.IsNullOrEmpty(dataRangeAttr))
            {
                chart.DataRange = sheet.Range[dataRangeAttr];
                chart.IsSeriesInRows = false;
            }
            else if (chartElement.Elements("series").Any())
            {
                var seriesList = chartElement.Elements("series").ToList();
                for (int s = 0; s < seriesList.Count; s++)
                {
                    var series = seriesList[s];
                    var name = series.Attribute("name")?.Value ?? $"Seri {s + 1}";
                    var points = series.Elements("point").ToList();
                    int row = chartStartRow + s;

                    sheet.Range[row, chartStartCol].Text = name;

                    for (int i = 0; i < points.Count; i++)
                    {
                        sheet.Range[chartStartRow - 1, chartStartCol + i + 1].Text = points[i].Attribute("label")?.Value;
                        sheet.Range[row, chartStartCol + i + 1].Number = double.Parse(points[i].Attribute("value")?.Value ?? "0");
                    }
                }

                int chartEndRow = chartStartRow + seriesList.Count - 1;
                int firstSeriesPoints = seriesList.FirstOrDefault()?.Elements("point")?.Count() ?? 0;
                int chartEndCol = chartStartCol + firstSeriesPoints;

                chart.DataRange = sheet.Range[chartStartRow - 1, chartStartCol, chartEndRow, chartEndCol];
                chart.IsSeriesInRows = true;
            }
            else
            {
                // 🔧 TABLODAN OTOMATİK VERİ AL
                var firstTable = sheetXml.Elements("table").FirstOrDefault();
                if (firstTable != null)
                {
                    string startCell = firstTable.Attribute("startCell")?.Value ?? "A1";
                    IRange range = sheet.Range[startCell];
                    int tableStartRow = range.Row;
                    int tableStartCol = range.Column;

                    int dataRows = firstTable.Elements("row").Count();
                    int dataCols = firstTable.Elements("row").Max(r => r.Elements("cell").Count());

                    int lastDataRow = tableStartRow + dataRows - 1;
                    int lastDataCol = tableStartCol + dataCols - 1;

                    chart.DataRange = sheet.Range[tableStartRow, tableStartCol, lastDataRow, lastDataCol];
                    chart.IsSeriesInRows = false;
                    //chart.DataLabels.IsValue = true;
                    chart.TopRow = chartStartRow;
                    chart.LeftColumn = chartStartCol;
                    chart.BottomRow = chartStartRow + 20;
                    chart.RightColumn = chartStartCol + 10;
                }
            }
        }





        private static ExcelChartType ParseChartType(string type)
        {
            return type.ToLower() switch
            {
                "line" => ExcelChartType.Line,
                "pie" => ExcelChartType.Pie,
                "bar" => ExcelChartType.Bar_Clustered,
                _ => ExcelChartType.Column_Clustered
            };
        }
    }
}
