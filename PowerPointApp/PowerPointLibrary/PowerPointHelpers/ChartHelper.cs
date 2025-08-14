using Syncfusion.OfficeChart;
using Syncfusion.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace PowerPointLibrary.PowerPointHelpers
{
    public static class ChartHelper
    {
        public static void AddChart(ISlide slide, XElement chartElement)
        {
            string? chartTypeStr = chartElement.Attribute("type")?.Value;
            if (string.IsNullOrWhiteSpace(chartTypeStr))
            {
                throw new Exception("Chart için 'type' niteliği zorunludur ve boş olamaz.");
            }

            double x = (double.TryParse(chartElement.Attribute("x")?.Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double dx) ? dx : 1) * 28.3465;
            double y = (double.TryParse(chartElement.Attribute("y")?.Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double dy) ? dy : 1) * 28.3465;
            double cx = (double.TryParse(chartElement.Attribute("w")?.Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double dw) ? dw : 15) * 28.3465;
            double cy = (double.TryParse(chartElement.Attribute("h")?.Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double dh) ? dh : 15) * 28.3465;

            switch (chartTypeStr.ToLower())
            {
                case "pie":
                    AddPieChart(slide, chartElement, x, y, cx, cy);
                    break;
                case "column":
                case "bar":
                    AddColumnChart(slide, chartElement, x, y, cx, cy);
                    break;
                case "line":
                    AddLineChart(slide, chartElement, x, y, cx, cy);
                    break;
                case "area":
                    AddAreaChart(slide, chartElement, x, y, cx, cy);
                    break;
                case "doughnut":
                    AddDoughnutChart(slide, chartElement, x, y, cx, cy);
                    break;
                case "scatter":
                    AddScatterChart(slide, chartElement, x, y, cx, cy);
                    break;
                default:
                    throw new Exception($"Desteklenmeyen chart tipi: {chartTypeStr}");
            }
        }

        public static void AddPieChart(ISlide slide, XElement chartElement, double x, double y, double cx, double cy)
        {
            IPresentationChart chart = slide.Charts.AddChart(x, y, cx, cy);
            chart.ChartType = OfficeChartType.Pie;

            string? title = chartElement.Attribute("title")?.Value;
            if (!string.IsNullOrEmpty(title))
            {
                chart.ChartTitle = title;
            }
            chart.Series.Clear();

            List<XElement> categories = chartElement.Elements("category").ToList();
            if (categories.Count == 0) return;

            for (int i = 0; i < categories.Count; i++)
            {
                chart.ChartData.SetValue(i + 2, 1, categories[i].Attribute("name")?.Value ?? $"Kategori {i + 1}");
            }

            IOfficeChartSerie serie = chart.Series.Add();
            serie.Name = chartElement.Attribute("seriesName")?.Value ?? "Veriler";

            for (int i = 0; i < categories.Count; i++)
            {
                double value = double.TryParse(categories[i].Attribute("value")?.Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double val) ? val : 0;
                chart.ChartData.SetValue(i + 2, 2, value);
            }

            serie.Values = chart.ChartData[2, 2, categories.Count + 1, 2];
            chart.PrimaryCategoryAxis.CategoryLabels = chart.ChartData[2, 1, categories.Count + 1, 1];
        }

        public static void AddColumnChart(ISlide slide, XElement chartElement, double x, double y, double cx, double cy)
        {
            IPresentationChart chart = slide.Charts.AddChart(x, y, cx, cy);
            chart.ChartType = OfficeChartType.Column_Clustered;

            string? title = chartElement.Attribute("title")?.Value;
            if (!string.IsNullOrEmpty(title))
            {
                chart.ChartTitle = title;
            }

            chart.Series.Clear();

            List<XElement> categories = chartElement.Elements("category").ToList();
            List<XElement> series = chartElement.Elements("series").ToList();

            if (categories.Count == 0 || series.Count == 0) return;

            for (int i = 0; i < categories.Count; i++)
            {
                chart.ChartData.SetValue(i + 2, 1, categories[i].Attribute("name")?.Value ?? $"Kategori {i + 1}");
            }

            for (int seriesIndex = 0; seriesIndex < series.Count; seriesIndex++)
            {
                XElement seriesElement = series[seriesIndex];
                IOfficeChartSerie serie = chart.Series.Add();
                serie.Name = seriesElement.Attribute("name")?.Value ?? $"Seri {seriesIndex + 1}";

                List<XElement> dataPoints = seriesElement.Elements("point").ToList();
                for (int i = 0; i < Math.Min(categories.Count, dataPoints.Count); i++)
                {
                    double value = double.TryParse(dataPoints[i].Attribute("value")?.Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double val) ? val : 0;
                    chart.ChartData.SetValue(i + 2, seriesIndex + 2, value);
                }

                serie.Values = chart.ChartData[2, seriesIndex + 2, categories.Count + 1, seriesIndex + 2];
            }

            chart.PrimaryCategoryAxis.CategoryLabels = chart.ChartData[2, 1, categories.Count + 1, 1];
        }

        public static void AddLineChart(ISlide slide, XElement chartElement, double x, double y, double cx, double cy)
        {
            IPresentationChart chart = slide.Charts.AddChart(x, y, cx, cy);
            chart.ChartType = OfficeChartType.Line;

            string? title = chartElement.Attribute("title")?.Value;
            if (!string.IsNullOrEmpty(title))
            {
                chart.ChartTitle = title;
            }

            chart.Series.Clear();

            List<XElement> categories = chartElement.Elements("category").ToList();
            List<XElement> series = chartElement.Elements("series").ToList();

            if (categories.Count == 0 || series.Count == 0) return;

            for (int i = 0; i < categories.Count; i++)
            {
                chart.ChartData.SetValue(i + 2, 1, categories[i].Attribute("name")?.Value ?? $"Kategori {i + 1}");
            }

            for (int seriesIndex = 0; seriesIndex < series.Count; seriesIndex++)
            {
                XElement seriesElement = series[seriesIndex];
                IOfficeChartSerie serie = chart.Series.Add();
                serie.Name = seriesElement.Attribute("name")?.Value ?? $"Seri {seriesIndex + 1}";

                List<XElement> dataPoints = seriesElement.Elements("point").ToList();
                for (int i = 0; i < Math.Min(categories.Count, dataPoints.Count); i++)
                {
                    double value = double.TryParse(dataPoints[i].Attribute("value")?.Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double val) ? val : 0;
                    chart.ChartData.SetValue(i + 2, seriesIndex + 2, value);
                }

                serie.Values = chart.ChartData[2, seriesIndex + 2, categories.Count + 1, seriesIndex + 2];
            }

            chart.PrimaryCategoryAxis.CategoryLabels = chart.ChartData[2, 1, categories.Count + 1, 1];
        }

        public static void AddAreaChart(ISlide slide, XElement chartElement, double x, double y, double cx, double cy)
        {
            IPresentationChart chart = slide.Charts.AddChart(x, y, cx, cy);
            chart.ChartType = OfficeChartType.Area;

            string? title = chartElement.Attribute("title")?.Value;
            if (!string.IsNullOrEmpty(title))
            {
                chart.ChartTitle = title;
            }

            chart.Series.Clear();

            List<XElement> categories = chartElement.Elements("category").ToList();
            List<XElement> series = chartElement.Elements("series").ToList();

            if (categories.Count == 0 || series.Count == 0) return;

            for (int i = 0; i < categories.Count; i++)
            {
                chart.ChartData.SetValue(i + 2, 1, categories[i].Attribute("name")?.Value ?? $"Kategori {i + 1}");
            }

            for (int seriesIndex = 0; seriesIndex < series.Count; seriesIndex++)
            {
                XElement seriesElement = series[seriesIndex];
                IOfficeChartSerie serie = chart.Series.Add();
                serie.Name = seriesElement.Attribute("name")?.Value ?? $"Seri {seriesIndex + 1}";

                List<XElement> dataPoints = seriesElement.Elements("point").ToList();
                for (int i = 0; i < Math.Min(categories.Count, dataPoints.Count); i++)
                {
                    double value = double.TryParse(dataPoints[i].Attribute("value")?.Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double val) ? val : 0;
                    chart.ChartData.SetValue(i + 2, seriesIndex + 2, value);
                }

                serie.Values = chart.ChartData[2, seriesIndex + 2, categories.Count + 1, seriesIndex + 2];
            }

            chart.PrimaryCategoryAxis.CategoryLabels = chart.ChartData[2, 1, categories.Count + 1, 1];
        }

        public static void AddDoughnutChart(ISlide slide, XElement chartElement, double x, double y, double cx, double cy)
        {
            IPresentationChart chart = slide.Charts.AddChart(x, y, cx, cy);
            chart.ChartType = OfficeChartType.Doughnut;

            string? title = chartElement.Attribute("title")?.Value;
            if (!string.IsNullOrEmpty(title))
            {
                chart.ChartTitle = title;
            }

            chart.Series.Clear();

            List<XElement> categories = chartElement.Elements("category").ToList();
            if (categories.Count == 0) return;

            for (int i = 0; i < categories.Count; i++)
            {
                chart.ChartData.SetValue(i + 2, 1, categories[i].Attribute("name")?.Value ?? $"Kategori {i + 1}");
            }

            IOfficeChartSerie serie = chart.Series.Add();
            serie.Name = chartElement.Attribute("seriesName")?.Value ?? "Veriler";

            for (int i = 0; i < categories.Count; i++)
            {
                double value = double.TryParse(categories[i].Attribute("value")?.Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double val) ? val : 0;
                chart.ChartData.SetValue(i + 2, 2, value);
            }

            serie.Values = chart.ChartData[2, 2, categories.Count + 1, 2];
            chart.PrimaryCategoryAxis.CategoryLabels = chart.ChartData[2, 1, categories.Count + 1, 1];
        }

        public static void AddScatterChart(ISlide slide, XElement chartElement, double x, double y, double cx, double cy)
        {
            IPresentationChart chart = slide.Charts.AddChart(x, y, cx, cy);
            chart.ChartType = OfficeChartType.Scatter_Markers;

            string? title = chartElement.Attribute("title")?.Value;
            if (!string.IsNullOrEmpty(title))
            {
                chart.ChartTitle = title;
            }

            chart.Series.Clear();

            List<XElement> series = chartElement.Elements("series").ToList();
            if (series.Count == 0) return;

            for (int seriesIndex = 0; seriesIndex < series.Count; seriesIndex++)
            {
                XElement seriesElement = series[seriesIndex];
                IOfficeChartSerie serie = chart.Series.Add();
                serie.Name = seriesElement.Attribute("name")?.Value ?? $"Seri {seriesIndex + 1}";

                List<XElement> dataPoints = seriesElement.Elements("point").ToList();
                for (int i = 0; i < dataPoints.Count; i++)
                {
                    double xValue = double.TryParse(dataPoints[i].Attribute("x")?.Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double xVal) ? xVal : 0;
                    double yValue = double.TryParse(dataPoints[i].Attribute("y")?.Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double yVal) ? yVal : 0;

                    chart.ChartData.SetValue(i + 2, seriesIndex * 2 + 1, xValue);
                    chart.ChartData.SetValue(i + 2, seriesIndex * 2 + 2, yValue);
                }

                serie.Values = chart.ChartData[2, seriesIndex * 2 + 2, dataPoints.Count + 1, seriesIndex * 2 + 2];
                serie.CategoryLabels = chart.ChartData[2, seriesIndex * 2 + 1, dataPoints.Count + 1, seriesIndex * 2 + 1];
            }
        }
    }
}
