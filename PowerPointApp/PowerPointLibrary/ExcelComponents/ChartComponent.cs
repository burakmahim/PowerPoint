using Syncfusion.XlsIO;
using System.Xml.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointLibrary.ExcelComponents
{
    public static class ChartComponent
    {
        public static void AddChart(XElement chartElement, IWorksheet sheet, IRange startRange, int rowCount, int colCount, int topRow, int chartWidth, int chartHeight)
        {
            string chartTypeStr = chartElement.Attribute("type")?.Value ?? "Column_Clustered";
            if (!Enum.TryParse(chartTypeStr, true, out ExcelChartType chartType))
            {
                chartType = ExcelChartType.Column_Clustered;
            }
            string? xAxis = chartElement.Attribute("xAxis")?.Value;
            string? yAxis = chartElement.Attribute("yAxis")?.Value;
            string? chartTitle = chartElement.Attribute("title")?.Value;
            string? dataRange = chartElement.Attribute("dataRange")?.Value;
            string? chartStartCell = chartElement.Attribute("startCell")?.Value;
            IChartShape chart = sheet.Charts.Add();
            chart.ChartType = chartType;

            if (dataRange != null)
            {
                chart.DataRange = sheet.Range[dataRange];
            }
            else
            {
                chart.DataRange = sheet.Range[
                startRange.Row,
                startRange.Column,
                startRange.Row + rowCount - 1,
                startRange.Column + colCount - 1];
            }
            chart.ChartTitle = chartTitle;
            chart.IsSeriesInRows = false;
            chart.PrimaryCategoryAxis.Title = xAxis;
            chart.PrimaryValueAxis.Title = yAxis;
            if (!string.IsNullOrEmpty(chartStartCell))
            {
                IRange chartStartRange = sheet.Range[chartStartCell];
                chart.TopRow = chartStartRange.Row;
                chart.LeftColumn = chartStartRange.Column;
            }
            else
            {
                chart.TopRow = topRow;
                chart.LeftColumn = startRange.Column;
            }
            chart.BottomRow = chart.TopRow + chartHeight;
            chart.RightColumn = chart.LeftColumn + chartWidth;
            //chart.TopRow = topRow;
            //chart.BottomRow = topRow + chartHeight;
            //chart.LeftColumn = startRange.Column;
            //chart.RightColumn = startRange.Column + chartWidth;
        }
    }
}