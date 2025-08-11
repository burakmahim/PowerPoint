//using System.Data;
//using System.IO;
//using System.Linq;
//using System.Xml.Linq;
//using System.Collections.Generic;
//using PowerPointLibrary.Exceptions;
//using Syncfusion.Pdf;
//using Syncfusion.XlsIO;



//#if NET48
//using Syncfusion.ExcelToPdfConverter;
//using Syncfusion.ExcelChartToImageConverter;

//#elif NET9_0
//using Syncfusion.XlsIORenderer;
//#endif


//namespace PowerPointLibrary.ExcelServices
//{
//    public static class ExcelChartImageGenerator
//    {
//        public static List<(string ChartName, byte[] ImageBytes)> ExtractChartImages(string xmlContent)
//        {
//            List<(string, byte[])> chartImages = new();

//            using ExcelEngine engine = new ExcelEngine();
//            IApplication app = engine.Excel;
//            app.DefaultVersion = ExcelVersion.Xlsx;

//            // Grafik desteği için Image converter
//            app.ChartToImageConverter = new ChartToImageConverter();

//            IWorkbook workbook = app.Workbooks.Open(new MemoryStream(ExcelLibrary.CreateExcelFromCustomXml(xmlContent)));

//            foreach (IWorksheet sheet in workbook.Worksheets)
//            {
//                foreach (IChart chart in sheet.Charts)
//                {
//                    using MemoryStream imgStream = new MemoryStream();
//                    chart.SaveAsImage(imgStream);
//                    chartImages.Add((chart.Name, imgStream.ToArray()));
//                }
//            }

//            return chartImages;
//        }
//    }
//}
