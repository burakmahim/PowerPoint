using Syncfusion.XlsIO;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;


namespace PowerPointLibrary.ExcelHelpers
{
    public class TableBuilder
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

                    if (DateTime.TryParse(value, out DateTime dateValue))
                    {
                        cell.DateTime = dateValue;
                        cell.NumberFormat = "dd.mm.yyyy";
                    }
                    else if (double.TryParse(value, out double numericValue))
                    {
                        cell.Number = numericValue;
                    }
                    else
                    {
                        cell.Text = value;
                    }

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

            XElement? conditionalFormats = table.Element("conditionalFormats");
            if (conditionalFormats != null)
            {
                ApplyConditionalFormatting(conditionalFormats, sheet, tableRange);
            }

            IListObject listObject = sheet.ListObjects.Create(tableName, tableRange);
            listObject.BuiltInTableStyle = TableBuiltInStyles.TableStyleMedium9;

            return startRange;
        }

        public static void ApplyConditionalFormatting(XElement conditionalFormats, IWorksheet sheet, IRange tableRange)
        {
            foreach (XElement conditionalFormat in conditionalFormats.Elements("conditionalFormat"))
            {
                string? cellRange       = conditionalFormat.Attribute("cellRange")?.Value ?? "";
                string? operatorValue   = conditionalFormat.Attribute("operator")?.Value ?? "";
                string? bgColor         = conditionalFormat.Attribute("bgColor")?.Value ?? "";
                string? type            = conditionalFormat.Attribute("type")?.Value ?? "";

                IRange targetRange;
                if (string.IsNullOrEmpty(cellRange))
                {
                    targetRange = tableRange;
                }
                else
                {
                    targetRange = sheet.Range[cellRange];
                }

                IConditionalFormats conditions = targetRange.ConditionalFormats;
                IConditionalFormat condition = conditions.AddCondition();

                string? firstFormula = conditionalFormat.Attribute("firstFormula")?.Value ?? "";
                string? secondFormula = conditionalFormat.Attribute("secondFormula")?.Value;

                switch (type)
                {
                    case "cellvalue":
                        condition.FormatType = ExcelCFType.CellValue;
                        switch (operatorValue.ToLower())
                        {
                            case "less":
                                condition.Operator = ExcelComparisonOperator.Less;
                                break;
                            case "greater":
                                condition.Operator = ExcelComparisonOperator.Greater;
                                break;
                            case "equal":
                                condition.Operator = ExcelComparisonOperator.Equal;
                                break;
                            case "notequal":
                                condition.Operator = ExcelComparisonOperator.NotEqual;
                                break;
                            case "lessorequal":
                                condition.Operator = ExcelComparisonOperator.LessOrEqual;
                                break;
                            case "greaterorequal":
                                condition.Operator = ExcelComparisonOperator.GreaterOrEqual;
                                break;
                            case "between":
                                condition.Operator = ExcelComparisonOperator.Between;
                                break;
                        }
                        condition.FirstFormula = firstFormula;
                        if (!string.IsNullOrEmpty(secondFormula))
                            condition.SecondFormula = secondFormula;
                        break;

                    case "formula":
                        condition.FormatType = ExcelCFType.Formula;
                        condition.FirstFormula = firstFormula;
                        break;

                    case "specifictext":
                        condition.FormatType = ExcelCFType.SpecificText;
                        switch (operatorValue.ToLower())
                        {
                            case "contains":
                                condition.Operator = ExcelComparisonOperator.ContainsText;
                                break;
                            case "notcontains":
                                condition.Operator = ExcelComparisonOperator.NotContainsText;
                                break;
                            case "beginswith":
                                condition.Operator = ExcelComparisonOperator.BeginsWith;
                                break;
                            case "endswith":
                                condition.Operator = ExcelComparisonOperator.EndsWith;
                                break;
                        }
                        condition.Text = firstFormula;
                        if (!string.IsNullOrEmpty(secondFormula))
                            condition.Text = secondFormula;
                        break;

                    case "timeperiod":
                        condition.FormatType = ExcelCFType.TimePeriod;
                        switch (operatorValue.ToLower())
                        {
                            case "yesterday":
                                condition.TimePeriodType = CFTimePeriods.Yesterday;
                                break;
                            case "today":
                                condition.TimePeriodType = CFTimePeriods.Today;
                                break;
                            case "tomorrow":
                                condition.TimePeriodType = CFTimePeriods.Tomorrow;
                                break;
                            case "last7days":
                                condition.TimePeriodType = CFTimePeriods.Last7Days;
                                break;
                            case "lastweek":
                                condition.TimePeriodType = CFTimePeriods.LastWeek;
                                break;
                            case "thisweek":
                                condition.TimePeriodType = CFTimePeriods.ThisWeek;
                                break;
                            case "nextweek":
                                condition.TimePeriodType = CFTimePeriods.NextWeek;
                                break;
                            case "lastmonth":
                                condition.TimePeriodType = CFTimePeriods.LastMonth;
                                break;
                            case "thismonth":
                                condition.TimePeriodType = CFTimePeriods.ThisMonth;
                                break;
                            case "nextmonth":
                                condition.TimePeriodType = CFTimePeriods.NextMonth;
                                break;
                        }
                        break;

                    case "duplicate":
                        condition.FormatType = ExcelCFType.Duplicate;
                        break;

                    case "unique":
                        condition.FormatType = ExcelCFType.Unique;
                        break;

                    case "blank":
                        condition.FormatType = ExcelCFType.Blank;
                        break;

                    default:
                        break;
                }

                if (!string.IsNullOrEmpty(bgColor))
                {
                    ExcelKnownColors? excelColor = GetExcelColor(bgColor);
                    if (excelColor.HasValue)
                    {
                        condition.BackColor = excelColor.Value;
                    }
                }

            }
        }

        public static ExcelKnownColors? GetExcelColor(string? colorName)
        {
            if (string.IsNullOrEmpty(colorName))
                return null;

            switch (colorName?.ToLower())
            {
                case "red": return ExcelKnownColors.Red;
                case "green": return ExcelKnownColors.Green;
                case "blue": return ExcelKnownColors.Blue;
                case "yellow": return ExcelKnownColors.Yellow;
                case "orange": return ExcelKnownColors.Orange;
                case "grey": return ExcelKnownColors.Grey_25_percent;
                case "lightblue": return ExcelKnownColors.Light_blue;
                case "lightgreen": return ExcelKnownColors.Light_green;
                case "pink": return ExcelKnownColors.Pink;
                default:
                    return null;
            }
        }
    }
}