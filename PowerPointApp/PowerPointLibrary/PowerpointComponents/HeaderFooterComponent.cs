using Syncfusion.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace PowerPointLibrary.PowerpointComponents
{
    public static class HeaderFooterComponent
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

        }
    }
}
