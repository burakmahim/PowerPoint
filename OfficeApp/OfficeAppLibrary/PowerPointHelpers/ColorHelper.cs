using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Syncfusion.Presentation;

namespace OfficeAppLibrary.PowerPointHelpers
{
    public static class ColorHelper
    {
        public static ColorObject ParseColor(string hexColor)
        {
            if (hexColor.StartsWith("#"))
                hexColor = hexColor.Substring(1);

            byte r = Convert.ToByte(hexColor.Substring(0, 2), 16);
            byte g = Convert.ToByte(hexColor.Substring(2, 2), 16);
            byte b = Convert.ToByte(hexColor.Substring(4, 2), 16);

            return (ColorObject)ColorObject.FromArgb(r, g, b);
        }
    }
}
