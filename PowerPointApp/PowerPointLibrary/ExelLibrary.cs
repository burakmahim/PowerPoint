using System;
using System.IO;
using System.Linq;

#if NET48 || NET9_0
using Syncfusion.XlsIO;
#endif

namespace PowerPointLibrary
{
    public static class ExcelLibrary
    {


        public static byte[] CreateExcelFromText(string rawText)
        {
        #if NET48 || NET9_0
            using ExcelEngine excelEngine = new ExcelEngine();
            IApplication application = excelEngine.Excel;
            application.DefaultVersion = ExcelVersion.Excel2016;

            IWorkbook workbook = application.Workbooks.Create(1);
            IWorksheet sheet = workbook.Worksheets[0];

            // Verileri satır-satır ve hücre-hücre yaz
            var lines = rawText.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < lines.Length; i++)
            {
                var cells = lines[i].Split(',');
                for (int j = 0; j < cells.Length; j++)
                {
                    sheet.Range[i + 1, j + 1].Text = cells[j].Trim();
                }
            }

            using MemoryStream ms = new MemoryStream();
            workbook.SaveAs(ms);
            return ms.ToArray();
        #else
            throw new PlatformNotSupportedException("Bu platform desteklenmiyor.");
        #endif
        }

    }
}
