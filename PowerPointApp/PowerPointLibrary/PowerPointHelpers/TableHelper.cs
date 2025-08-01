using Syncfusion.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace PowerPointLibrary.PowerPointHelpers
{
    public static class TableHelper
    {
        public static void AddTable(ISlide slide, XElement tableElement)
        {
            List<XElement> rows = tableElement.Elements("tr").ToList();
            if (rows.Count == 0) return;

            double x = (double.TryParse(tableElement.Attribute("x")?.Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double dx) ? dx : 2) * 28.3465;
            double y = (double.TryParse(tableElement.Attribute("y")?.Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double dy) ? dy : 5) * 28.3465;
            double cx = (double.TryParse(tableElement.Attribute("w")?.Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double dw) ? dw : 30) * 28.3465;
            double cy = (double.TryParse(tableElement.Attribute("h")?.Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double dh) ? dh : 3) * 28.3465;

            int rowCount = rows.Count;
            int colCount = rows[0].Elements("td").Count();

            ITable table = slide.Shapes.AddTable(rowCount,colCount, x, y, cx, cy);

            for (int r = 0; r < rowCount; r++)
            {
                List<XElement> cells = rows[r].Elements("td").ToList();
                for (int c = 0; c < colCount; c++)
                {
                    table.Rows[r].Cells[c].TextBody.AddParagraph(c < cells.Count ? cells[c].Value : "");
                }
            }
        }
    }
}
