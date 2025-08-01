using Syncfusion.XlsIO;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace PowerPointLibrary.ExcelHelpers
{
    public static class TableBuilder
    {
        public static IRange AddTable(XElement table, IWorksheet sheet, out int rowCount, out int colCount)
        {
            string? startCell = table.Attribute("startCell")?.Value ?? "A1";
            IRange startRange = sheet.Range[startCell];

            List<XElement> rows = table.Elements("row").ToList();
            rowCount = rows.Count;
            colCount = rows.Max(r => r.Elements("cell").Count());

            for (int rowIndex = 0; rowIndex < rowCount; rowIndex++)
            {
                List<XElement> cells = rows[rowIndex].Elements("cell").ToList();
                for (int colIndex = 0; colIndex < cells.Count; colIndex++)
                {
                    IRange cell = sheet[startRange.Row + rowIndex, startRange.Column + colIndex];
                    string? value = cells[colIndex].Value;
                    string? formula = cells[colIndex].Attribute("formula")?.Value;

                    if (double.TryParse(value, out double numericValue))
                        cell.Number = numericValue;
                    else
                        cell.Text = value;

                    if (!string.IsNullOrEmpty(formula))
                        cell.Formula = formula;

                    if (cells[colIndex].Attribute("bold")?.Value == "true")
                        cell.CellStyle.Font.Bold = true;
                }
            }

            sheet.Calculate();

            string? tableName = $"Table_{sheet.Name}_{sheet.ListObjects.Count + 1}";

            IRange tableRange = sheet.Range[
                startRange.Row,
                startRange.Column,
                startRange.Row + rowCount - 1,
                startRange.Column + colCount - 1];

            IListObject listObject = sheet.ListObjects.Create(tableName, tableRange);
            listObject.BuiltInTableStyle = TableBuiltInStyles.TableStyleMedium9;

            return startRange;
        }
    }
}
