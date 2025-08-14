using Syncfusion.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;


namespace PowerPointLibrary.PowerpointComponents
{
    public static class ListComponent
    {
        public static void AddList(XElement listElement, ISlide slide)
        {

            (double x, double y, double cx, double cy) = CoordinatesParser.CoordinateParser(listElement,1, 1, 5, 5);

            bool bold = bool.TryParse(listElement.Attribute("bold")?.Value, out bool b) && b;
            bool italic = bool.TryParse(listElement.Attribute("italic")?.Value, out bool i) && i;

            string listFontFamily = listElement.Attribute("font")?.Value ?? "Calibri";
            string listTextColor = listElement.Attribute("color")?.Value ?? "#000000";
            string listFontSize = listElement.Attribute("size")?.Value ?? "14";
            string listType = listElement.Attribute("type")?.Value ?? "bulleted";

            ListType listTypeParsed = Enum.TryParse<ListType>(listType, true, out ListType result) ? result : ListType.Bulleted;

            int listFontSizeParsed = int.TryParse(listFontSize, out int s) ? s : 14;

            IShape listBox = slide.AddTextBox(x, y, cx, cy);

            foreach (XElement item in listElement.Elements("item"))
            {
                IParagraph paragraph = listBox.TextBody.AddParagraph(item.Value);
                paragraph.ListFormat.Type = listTypeParsed;

                string itemFontFamily = item.Attribute("fontFamily")?.Value ?? listFontFamily;
                string itemTextColor = item.Attribute("color")?.Value ?? listTextColor;
                string itemFontSize = item.Attribute("fontSize")?.Value ?? listFontSize;

                int itemFontSizeParsed = int.Parse(itemFontSize);

                paragraph.Font.FontName = itemFontFamily;
                paragraph.Font.FontSize = itemFontSizeParsed;
                paragraph.Font.Color = ColorHelper.ParseColor(itemTextColor);
            }

        }

    }
}
