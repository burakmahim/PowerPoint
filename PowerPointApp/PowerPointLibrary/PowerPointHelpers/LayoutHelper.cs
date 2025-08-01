using Syncfusion.Presentation;
using System.Xml.Linq;

namespace PowerPointLibrary.PowerPointHelpers
{
    public static class LayoutHelper
    {
        public static void SetLayoutContent(ISlide slide, XElement slideElement, SlideLayoutType layoutType)
        {
            string? title = slideElement.Element("title")?.Value;
            string? subtitle = slideElement.Element("subtitle")?.Value;
            string? content = slideElement.Element("content")?.Value;
            string? leftcontent = slideElement.Element("leftcontent")?.Value;
            string? rightcontent = slideElement.Element("rightcontent")?.Value;

            int bodyCounter = 0;

            foreach (IShape shape in slide.Shapes)
            {
                switch (shape.PlaceholderFormat.Type)
                {
                    case PlaceholderType.Title:
                        if (!string.IsNullOrEmpty(title))
                            shape.TextBody.AddParagraph(title);
                        break;

                    case PlaceholderType.Subtitle:
                        if (!string.IsNullOrEmpty(subtitle))
                            shape.TextBody.AddParagraph(subtitle);
                        break;

                    case PlaceholderType.Body:
                        if (layoutType == SlideLayoutType.TwoContent)
                        {
                            if (bodyCounter == 0 && !string.IsNullOrEmpty(leftcontent))
                                shape.TextBody.AddParagraph(leftcontent);
                            else if (bodyCounter == 1 && !string.IsNullOrEmpty(rightcontent))
                                shape.TextBody.AddParagraph(rightcontent);

                            bodyCounter++;
                        }
                        else if (!string.IsNullOrEmpty(content))
                        {
                            shape.TextBody.AddParagraph(content);
                        }
                        break;

                    default:
                        if (!string.IsNullOrEmpty(content))
                            shape.TextBody.AddParagraph(content);
                        break;
                }
            }
        }
    }
}
