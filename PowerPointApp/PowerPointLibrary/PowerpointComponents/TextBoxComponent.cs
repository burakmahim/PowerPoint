using Syncfusion.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace PowerPointLibrary.PowerpointComponents
{
    class TextBoxComponent
    {
        public static void AddTextBox(XElement textboxElement, ISlide slide)
        {
            string? fontFamily = textboxElement.Attribute("fontFamily")?.Value;

            (double x, double y, double cx, double cy) = CoordinatesParser.CoordinateParser(textboxElement, 1, 1, 5, 5);

            string? text = textboxElement.Value;

            bool bold = bool.TryParse(textboxElement.Attribute("bold")?.Value, out bool b) && b;
            bool italic = bool.TryParse(textboxElement.Attribute("italic")?.Value, out bool i) && i;

            string? textColor = textboxElement.Attribute("textColor")?.Value ?? "#000000";
            string? backgroundColor = textboxElement.Attribute("backgroundColor")?.Value;
            int fontSize = int.TryParse(textboxElement.Attribute("fontSize")?.Value, out int fs) ? fs : 12;

            HorizontalAlignmentType alignment = Enum.TryParse(textboxElement.Attribute("alignment")?.Value, true, out HorizontalAlignmentType align) ? align : HorizontalAlignmentType.Left;

            IShape textbox = slide.AddTextBox(x, y, cx, cy);
            textbox.Fill.FillType = FillType.None;

            IParagraph paragraph = textbox.TextBody.AddParagraph(text);
            paragraph.HorizontalAlignment = alignment;
            paragraph.Font.Bold = bold;
            paragraph.Font.Italic = italic;
            paragraph.Font.FontSize = fontSize;
            paragraph.Font.FontName = fontFamily;
            paragraph.HorizontalAlignment = alignment;
            paragraph.Font.Color = ColorHelper.ParseColor(textColor);

            //shape.LineFormat.Fill.FillType = FillType.None;

            if (backgroundColor != null)
            {
                textbox.Fill.FillType = FillType.Solid;
                textbox.Fill.SolidFill.Color = ColorHelper.ParseColor(backgroundColor);
            }

        }
    }
}