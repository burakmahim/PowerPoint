using System.Data;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using System.Collections.Generic;
using PowerPointLibrary.Exceptions;
using Syncfusion.Pdf;
using Syncfusion.XlsIO;



#if NET48
using Syncfusion.ExcelToPdfConverter;
using Syncfusion.ExcelChartToImageConverter;

#elif NET9_0
using Syncfusion.XlsIORenderer;
#endif


namespace PowerPointLibrary.ExcelServices
{
    public static class ExcelTableParser
    {
        // 📌 ExcelSheet ve tablo haritası ile tüm tabloları alır
        public static Dictionary<string, DataTable> ParseAllTables(IWorksheet sheet, Dictionary<string, (int row, int col, int rowCount, int colCount)> tableMap)
        {
            var result = new Dictionary<string, DataTable>();

            foreach (var entry in tableMap)
            {
                string tableName = entry.Key;
                var (startRow, startCol, rowCount, colCount) = entry.Value;

                // İlgili hücre aralığı
                IRange range = sheet.Range[startRow, startCol, startRow + rowCount - 1, startCol + colCount - 1];

                DataTable dt = ConvertRangeToDataTable(range);
                result[tableName] = dt;
            }

            return result;
        }

        // 📊 Belirli IRange aralığını DataTable olarak döner (formül sonuçları dahil)
        public static DataTable ConvertRangeToDataTable(IRange range)
        {
            DataTable dt = new DataTable();

            int rowCount = range.LastRow - range.Row + 1;
            int colCount = range.LastColumn - range.Column + 1;

            // Sütun başlıkları ekleniyor
            for (int col = 0; col < colCount; col++)
                dt.Columns.Add($"Sütun {col + 1}");

            // Satır verileri ekleniyor
            for (int row = 0; row < rowCount; row++)
            {
                DataRow dr = dt.NewRow();
                for (int col = 0; col < colCount; col++)
                {
                    // DisplayText: hem değer hem de formül sonucu için güvenlidir
                    dr[col] = range[row + 1, col + 1].DisplayText;
                }
                dt.Rows.Add(dr);
            }

            return dt;
        }

        public static byte[] ConvertChartOnlyToPdf(string xmlContent)
        {
#if NET48
            var workbook = ExcelParserService.CreateChartOnlyWorkbookFromXml(xmlContent);

            ExcelToPdfConverter converter = new ExcelToPdfConverter(workbook);
            converter.ChartToImageConverter = new ChartToImageConverter();

            ExcelToPdfConverterSettings settings = new ExcelToPdfConverterSettings
            {
                LayoutOptions = LayoutOptions.FitSheetOnOnePage
            };

            PdfDocument pdfDocument = converter.Convert(settings);

            using MemoryStream outputStream = new MemoryStream();
            pdfDocument.Save(outputStream);
            return outputStream.ToArray();

#elif NET9_0
    var workbook = ExcelParserService.CreateChartOnlyWorkbookFromXml(xmlContent);

    XlsIORendererSettings settings = new XlsIORendererSettings
    {
        LayoutOptions = LayoutOptions.FitSheetOnOnePage
    };

    XlsIORenderer renderer = new XlsIORenderer();
    PdfDocument pdfDocument = renderer.ConvertToPDF(workbook, settings);

    using MemoryStream outputStream = new MemoryStream();
    pdfDocument.Save(outputStream);
    return outputStream.ToArray();

#else
    throw new PlatformNotSupportedException("Bu platform desteklenmiyor.");
#endif
        }
    }
}
