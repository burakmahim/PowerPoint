using Syncfusion.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace OfficeAppLibrary.PowerPointHelpers
{
    public static class HeaderFooterHelper
    {
        public static void SetHeaderFooter(XElement document, ISlide slide)
        {
            XElement? footerElement = document.Element("footer");
            if (footerElement == null) return;

            bool enableFooter = bool.TryParse(document.Element("settings")?.Attribute("footer")?.Value, out bool result) && result;
            string? footerText = footerElement.Value;

            slide.HeadersFooters.Footer.Visible = enableFooter;
            slide.HeadersFooters.Footer.Text = footerText;

            bool enableSlideNumber = bool.TryParse(document.Element("settings")?.Attribute("slideNumber")?.Value, out bool resultNumber) && resultNumber;

            slide.HeadersFooters.SlideNumber.Visible = enableSlideNumber;

            double x = (double.TryParse(footerElement.Attribute("x")?.Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double dx) ? dx : 1) * 28.3465;
            double y = (double.TryParse(footerElement.Attribute("y")?.Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double dy) ? dy : 12.5) * 28.3465;
            double cx = (double.TryParse(footerElement.Attribute("w")?.Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double dw) ? dw : 23) * 28.3465;
            double cy = (double.TryParse(footerElement.Attribute("h")?.Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double dh) ? dh : 1.2) * 28.3465;

        }
    }
}
