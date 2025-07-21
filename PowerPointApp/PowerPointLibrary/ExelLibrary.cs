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
        public static byte[] CreateExcelFromText(string rawInput)
        {
        #if NET48 || NET9_0
            if (string.IsNullOrWhiteSpace(rawInput))
                throw new ArgumentException("Boş veri gönderildi.");

            using ExcelEngine excelEngine = new ExcelEngine();
            IApplication application = excelEngine.Excel;
            application.DefaultVersion = ExcelVersion.Excel2016;

            IWorkbook workbook = application.Workbooks.Create(1);
            IWorksheet sheet = workbook.Worksheets[0];

            // Satırları al
            string[] lines = rawInput.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

            for (int row = 0; row < lines.Length; row++)
            {
                string[] cells = lines[row].Split(',');

                for (int col = 0; col < cells.Length; col++)
                {
                    sheet.Range[row + 1, col + 1].Text = cells[col].Trim();
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
