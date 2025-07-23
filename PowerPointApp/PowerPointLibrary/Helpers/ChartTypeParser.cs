using Syncfusion.XlsIO;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointLibrary.Helpers
{
    public static class ChartTypeParser
    {
        public static ExcelChartType Parse(string type)
        {
            return type.ToLowerInvariant() switch
            {
                "line" => ExcelChartType.Line,
                "pie" => ExcelChartType.Pie,
                "bar" => ExcelChartType.Bar_Clustered,
                "doughnut" => ExcelChartType.Doughnut,
                "area" => ExcelChartType.Area,
                "scatter" => ExcelChartType.Scatter_Markers,
                "bubble" => ExcelChartType.Bubble,
                _ => ExcelChartType.Column_Clustered // varsayılan
            };
        }
    }
}
