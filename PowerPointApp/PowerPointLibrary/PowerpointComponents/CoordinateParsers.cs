using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace PowerPointLibrary.PowerpointComponents
{
    public static class CoordinatesParser
    {
        public static (double x, double y, double cx, double cy) CoordinateParser(XElement element, double defaultX, double defaultY, double defaultCx, double defaultCy)
        {
            const double constant = 28.3465;

            double x = (double.TryParse(element.Attribute("x")?.Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double dx) ? dx : defaultX) * constant;
            double y = (double.TryParse(element.Attribute("y")?.Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double dy) ? dy : defaultY) * constant;
            double cx = (double.TryParse(element.Attribute("w")?.Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double dw) ? dw : defaultCx) * constant;
            double cy = (double.TryParse(element.Attribute("h")?.Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double dh) ? dh : defaultCy) * constant;

            return (x, y, cx, cy);
        }
    }
}
