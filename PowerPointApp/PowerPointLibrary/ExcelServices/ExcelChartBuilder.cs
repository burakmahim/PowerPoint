using Syncfusion.XlsIO;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace PowerPointLibrary.ExcelServices
{
    public static class ExcelChartBuilder
    {
        public static int AddChart(XElement chartElement, IWorksheet sheet, int defaultStartRow,
            Dictionary<string, (int StartRow, int StartCol, int RowCount, int ColCount)> tableMap)
        {
            int chartStartCol = int.TryParse(chartElement.Attribute("startCol")?.Value, out int c) ? c : 1;
            int chartStartRow = int.TryParse(chartElement.Attribute("startRow")?.Value, out int r) ? r : defaultStartRow;
            int chartWidth = int.TryParse(chartElement.Attribute("chartWidth")?.Value, out int w) ? w : 10;
            int chartHeight = int.TryParse(chartElement.Attribute("chartHeight")?.Value, out int h) ? h : 15;

            var chart = sheet.Charts.Add();
            chart.ChartType = Enum.TryParse(chartElement.Attribute("type")?.Value, true, out ExcelChartType chartType)
                ? chartType
                : ExcelChartType.Column_Clustered;

            chart.ChartTitle = chartElement.Attribute("title")?.Value ?? "";
            chart.PrimaryCategoryAxis.Title = chartElement.Attribute("xAxis")?.Value ?? "";
            chart.PrimaryValueAxis.Title = chartElement.Attribute("yAxis")?.Value ?? "";

            bool dataRangeSet = false;

            var sourceTableName = chartElement.Attribute("sourceTable")?.Value;
            if (!string.IsNullOrWhiteSpace(sourceTableName) && tableMap.TryGetValue(sourceTableName, out var tableInfo))
            {
                chart.DataRange = sheet.Range[tableInfo.StartRow, tableInfo.StartCol,
                                              tableInfo.StartRow + tableInfo.RowCount - 1,
                                              tableInfo.StartCol + tableInfo.ColCount - 1];
                chart.IsSeriesInRows = false;
                dataRangeSet = true;
            }

            // 2️⃣ dataRange varsa (manuel hücre aralığı)
            var manualRange = chartElement.Attribute("dataRange")?.Value;
            if (!dataRangeSet && !string.IsNullOrWhiteSpace(manualRange))
            {
                chart.DataRange = sheet.Range[manualRange];
                chart.IsSeriesInRows = false;
                dataRangeSet = true;
            }

            // 3️⃣ series ile manuel veri varsa
            var seriesElements = chartElement.Elements("series").ToList();
            if (!dataRangeSet && seriesElements.Any())
            {
                int seriesRowStart = chartStartRow - seriesElements.Count - 2;
                int labelRow = seriesRowStart - 1;

                int maxPoints = 0;

                for (int s = 0; s < seriesElements.Count; s++)
                {
                    var series = seriesElements[s];
                    var points = series.Elements("point").ToList();
                    maxPoints = Math.Max(maxPoints, points.Count);

                    sheet.Range[seriesRowStart + s, chartStartCol].Text = series.Attribute("name")?.Value ?? $"Seri {s + 1}";

                    for (int i = 0; i < points.Count; i++)
                    {
                        sheet.Range[labelRow, chartStartCol + i + 1].Text = points[i].Attribute("label")?.Value;
                        sheet.Range[seriesRowStart + s, chartStartCol + i + 1].Number = double.Parse(points[i].Attribute("value")?.Value ?? "0");
                    }
                }

                chart.DataRange = sheet.Range[labelRow, chartStartCol,
                                              seriesRowStart + seriesElements.Count - 1,
                                              chartStartCol + maxPoints];
                chart.IsSeriesInRows = true;
                dataRangeSet = true;
            }

            chart.TopRow = chartStartRow;
            chart.LeftColumn = chartStartCol;
            chart.BottomRow = chartStartRow + chartHeight;
            chart.RightColumn = chartStartCol + chartWidth;

            return chart.BottomRow + 2;
        }
    }
}
