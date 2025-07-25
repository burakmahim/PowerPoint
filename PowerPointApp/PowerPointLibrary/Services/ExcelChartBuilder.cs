using Syncfusion.XlsIO;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using PowerPointLibrary.Helpers;


namespace PowerPointLibrary.Services
{
    public static class ExcelChartBuilder
    {


        public static void AddChart(XElement chartElement, IWorksheet sheet, int chartStartRow,
            Dictionary<string, (int StartRow, int StartCol, int RowCount, int ColCount)> tableMap)
        {
            int chartStartCol = 1;

            var chart = sheet.Charts.Add();
            chart.ChartType = ChartTypeParser.Parse(chartElement.Attribute("type")?.Value ?? "Column");
            chart.ChartTitle = chartElement.Attribute("title")?.Value ?? "";
            chart.PrimaryCategoryAxis.Title = chartElement.Attribute("xAxis")?.Value ?? "";
            chart.PrimaryValueAxis.Title = chartElement.Attribute("yAxis")?.Value ?? "";

            // 1️⃣ sourceTable ile veri seçimi 
            var sourceTableName = chartElement.Attribute("sourceTable")?.Value;
            if (!string.IsNullOrWhiteSpace(sourceTableName) && tableMap.ContainsKey(sourceTableName))
            {
                var info = tableMap[sourceTableName];
                int startRow = info.StartRow;
                int startCol = info.StartCol;
                int endRow = startRow + info.RowCount - 1;
                int endCol = startCol + info.ColCount - 1;

                chart.DataRange = sheet.Range[startRow, startCol, endRow, endCol];
                chart.IsSeriesInRows = false;
            }
            // 2️⃣ dataRange ile manuel aralık seçimi
            if (!string.IsNullOrEmpty(chartElement.Attribute("dataRange")?.Value))
            {
                chart.DataRange = sheet.Range[chartElement.Attribute("dataRange")?.Value];
                chart.IsSeriesInRows = false;
            }
            // 3️⃣ series ile manuel veri girişi
            if (chartElement.Elements("series").Any())
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

            // 📐 Grafik konumu
            chart.TopRow = chartStartRow;
            chart.LeftColumn = chartStartCol;
            chart.BottomRow = chartStartRow + 20;
            chart.RightColumn = chartStartCol + 10;
        }


    }
}
