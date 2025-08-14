using Syncfusion.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;


namespace PowerPointLibrary.PowerPointHelpers
{
    public static class ShapeHelper
    {
        public static void AddShape(XElement shapeElement, ISlide slide)
        {
            string? fontFamily = shapeElement.Attribute("fontFamily")?.Value;

            double x = (double.TryParse(shapeElement.Attribute("x")?.Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double dx) ? dx : 1) * 28.3465;
            double y = (double.TryParse(shapeElement.Attribute("y")?.Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double dy) ? dy : 1) * 28.3465;
            double cx = (double.TryParse(shapeElement.Attribute("w")?.Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double dw) ? dw : 5) * 28.3465;
            double cy = (double.TryParse(shapeElement.Attribute("h")?.Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double dh) ? dh : 5) * 28.3465;

            string? text = shapeElement.Value;

            AutoShapeType shapeType = Enum.TryParse<AutoShapeType>(shapeElement.Attribute("shapeType")?.Value ?? "Rectangle", true, out AutoShapeType st) ? st : AutoShapeType.Rectangle;

            bool bold = bool.TryParse(shapeElement.Attribute("bold")?.Value, out bool b) && b;
            bool italic = bool.TryParse(shapeElement.Attribute("italic")?.Value, out bool i) && i;

            string? textColor = shapeElement.Attribute("textColor")?.Value ?? "#000000";
            string? backgroundColor = shapeElement.Attribute("backgroundColor")?.Value;
            int fontSize = int.TryParse(shapeElement.Attribute("fontSize")?.Value, out int fs) ? fs : 12;

            HorizontalAlignmentType alignment = Enum.TryParse(shapeElement.Attribute("alignment")?.Value, true, out HorizontalAlignmentType align) ? align : HorizontalAlignmentType.Left;

            IShape shape = slide.Shapes.AddShape(shapeType, x, y, cx, cy);
            shape.Fill.FillType = FillType.None;

            IParagraph paragraph = shape.TextBody.AddParagraph(text);
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
                shape.Fill.FillType = FillType.Solid;
                shape.Fill.SolidFill.Color = ColorHelper.ParseColor(backgroundColor);
            }
        }
    }
}
