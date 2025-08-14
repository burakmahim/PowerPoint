using Syncfusion.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace PowerPointLibrary.PowerpointComponents
{
    public static class TableComponent
    {
        public static void AddTable(ISlide slide, XElement tableElement)
        {
            List<XElement> rows = tableElement.Elements("tr").ToList();
            if (rows.Count == 0) return;

            (double x, double y, double cx, double cy) = CoordinatesParser.CoordinateParser(tableElement, 1, 1, 5, 5);

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
