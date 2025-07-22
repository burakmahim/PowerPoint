using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;


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

            IWorkbook workbook = application.Workbooks.Create(1);
            IWorksheet sheet = workbook.Worksheets[0];

            var doc = XDocument.Parse(xmlContent);
            var sheetElement = doc.Root?.Element("sheet");
            if (sheetElement == null)
                throw new InvalidDataException("Sheet bulunamadı.");

            var sheetName = sheetElement.Attribute("name")?.Value ?? "Sayfa1";
            sheet.Name = sheetName;

            // 📌 Varsayılan başlangıç konumları
            int tableStartRow = 1, tableStartCol = 1;
            int chartStartRow = -1, chartStartCol = -1;

            // 📌 Tabloyu işle
            var table = sheetElement.Element("table");
            var rows = table?.Elements("row").ToList();
            if (table != null && rows != null)
            {
                var tableStartCell = table.Attribute("startCell")?.Value;
                if (!string.IsNullOrEmpty(tableStartCell))
                    (tableStartRow, tableStartCol) = GetRowColFromCell(tableStartCell);

                for (int i = 0; i < rows.Count; i++)
                {
                    var cells = rows[i].Elements("cell").ToList();
                    for (int j = 0; j < cells.Count; j++)
                    {
                        var cell = sheet.Range[tableStartRow + i, tableStartCol + j];
                        cell.Text = cells[j].Value;
                        if (cells[j].Attribute("bold")?.Value == "true")
                            cell.CellStyle.Font.Bold = true;
                    }
                }

                // 📌 Chart için başlangıç ayarlanmadıysa tablonun altına koy
                if (chartStartRow == -1)
                {
                    chartStartRow = tableStartRow + rows.Count + 2;
                    chartStartCol = tableStartCol;
                }
            }

            // 📊 Chart varsa işle
            var chartElement = sheetElement.Element("chart");
            if (chartElement != null)
            {
                var chartStartCell = chartElement.Attribute("startCell")?.Value;
                if (!string.IsNullOrEmpty(chartStartCell))
                    (chartStartRow, chartStartCol) = GetRowColFromCell(chartStartCell);

                var chart = sheet.Charts.Add();
                chart.ChartType = ParseChartType(chartElement.Attribute("type")?.Value ?? "Column");
                chart.ChartTitle = chartElement.Attribute("title")?.Value ?? "";

                chart.PrimaryCategoryAxis.Title = chartElement.Attribute("xAxis")?.Value ?? "";
                chart.PrimaryValueAxis.Title = chartElement.Attribute("yAxis")?.Value ?? "";

                int seriesCount = 0;
                var seriesList = chartElement.Elements("series").ToList();

                foreach (var series in seriesList)
                {
                    string seriesName = series.Attribute("name")?.Value ?? $"Seri {seriesCount + 1}";
                    var points = series.Elements("point").ToList();
                    var dataRow = chartStartRow + seriesCount;

                    sheet.Range[dataRow, chartStartCol].Text = seriesName;

                    for (int i = 0; i < points.Count; i++)
                    {
                        sheet.Range[chartStartRow - 1, chartStartCol + i + 1].Text = points[i].Attribute("label")?.Value ?? $"Etiket {i + 1}";
                        sheet.Range[dataRow, chartStartCol + i + 1].Number = double.Parse(points[i].Attribute("value")?.Value ?? "0");
                    }

                    seriesCount++;
                }

                int chartLastRow = chartStartRow + seriesCount - 1;
                int chartLastCol = chartStartCol + seriesList.First().Elements("point").Count();

                chart.DataRange = sheet.Range[chartStartRow - 1, chartStartCol, chartLastRow, chartLastCol];
                chart.IsSeriesInRows = true;
            }

            using MemoryStream ms = new MemoryStream();
            workbook.SaveAs(ms);
            return ms.ToArray();
#else
    throw new PlatformNotSupportedException("Bu platform desteklenmiyor.");
#endif
        }


        private static (int Row, int Col) GetRowColFromCell(string cell)
        {
            // Örn: A1 => (1, 1), B2 => (2, 2)
            int col = 0;
            int rowIndex = 0;
            foreach (char ch in cell)
            {
                if (char.IsLetter(ch))
                {
                    col = col * 26 + (char.ToUpper(ch) - 'A' + 1);
                }
                else if (char.IsDigit(ch))
                {
                    rowIndex = rowIndex * 10 + (ch - '0');
                }
            }
            return (rowIndex, col);
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

