using Syncfusion.XlsIO;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace PowerPointLibrary.Services
{
    public static class ExcelTableBuilder
    {
        public static int AddTable(XElement table, IWorksheet sheet, int startRow, int startCol, Dictionary<string, (int, int, int, int)> tableMap)
        {
            string? startCell = table.Attribute("startCell")?.Value;
            if (!string.IsNullOrWhiteSpace(startCell))
            {
                IRange range = sheet.Range[startCell];
                startRow = range.Row;
                startCol = range.Column;
            }

            List<XElement> rows = table.Elements("row").ToList();

            for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++)
            {
                List<XElement> cells = rows[rowIndex].Elements("cell").ToList();
                for (int colIndex = 0; colIndex < cells.Count; colIndex++)
                {
                    IRange cell = sheet[startRow + rowIndex, startCol + colIndex];
                    XElement cellElement = cells[colIndex];
                    string value = cellElement.Value?.Trim() ?? "";
                    string? formula = cellElement.Attribute("formula")?.Value;


                    if (double.TryParse(value, out double numericValue))
                        cell.Number = numericValue;
                    else
                        cell.Text = value;

                    if (!string.IsNullOrEmpty(formula))
                    {
                        cell.Formula = formula;
                    }

                    if (cells[colIndex].Attribute("bold")?.Value == "true")
                        cell.CellStyle.Font.Bold = true;
                }
            }

            // 📌 Tablo bilgisi map'e ekleniyor
            string? name = table.Attribute("name")?.Value;
            if (!string.IsNullOrWhiteSpace(name))
            {
                int rowCount = rows.Count;
                int colCount = rows.Max(r => r.Elements("cell").Count());
                tableMap[name] = (startRow, startCol, rowCount, colCount);
            }

            // 📊 Formüller hesaplanır
            sheet.EnableSheetCalculations();
            sheet.Calculate();

            sheet.UsedRange.AutofitColumns();
            return startRow + rows.Count;
        }
    }
}

