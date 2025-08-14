using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using System.Collections.Generic;
using PowerPointLibrary.Exceptions;
using Syncfusion.Pdf;
using Syncfusion.XlsIO;
using Syncfusion.Presentation;
using Syncfusion.XlsIO.Parser.Biff_Records;
using PowerPointLibrary.ExcelComponents;


#if NET48
using Syncfusion.ExcelToPdfConverter;
using Syncfusion.ExcelChartToImageConverter;
#elif NET9_0
using Syncfusion.XlsIORenderer;
#endif

namespace PowerPointLibrary
{
    public static class ExcelLibrary
    {
        public class TableInfo
        {
            public IRange Range { get; set; }
            public int RowCount { get; set; }
            public int ColCount { get; set; }
        }
            public static byte[] CreateExcelFromCustomXml(string xmlContent)
            {
                try
                {
                    using ExcelEngine excelEngine = new ExcelEngine();
                    IApplication application = excelEngine.Excel;
                    application.DefaultVersion = ExcelVersion.Xlsx;

                    IWorkbook workbook = application.Workbooks.Create(0);

                    int pageCounter = 1;
                    XElement document = XElement.Parse(xmlContent);

                    foreach (XElement sheetXml in document.Descendants("sheet"))
                    {
                        string? sheetName = sheetXml.Attribute("name")?.Value ?? $"Sayfa{pageCounter}";
                        IWorksheet sheet = workbook.Worksheets.Create(sheetName);
                        sheet.EnableSheetCalculations();

                        int currentRow = 1;
                        Dictionary<string, TableInfo> tables = new Dictionary<string, TableInfo>();

                        TableInfo? lastTable = null;

                        foreach (XElement element in sheetXml.Elements())
                        {
                            if (element.Name == "table")
                            {
                                if (element.Attribute("startCell") == null)
                                {
                                    element.SetAttributeValue("startCell", $"A{currentRow}");
                                }

                                IRange tableRange = TableComponent.AddTable(element, sheet, out int rowCount, out int colCount);

                                TableInfo tableInfo = new TableInfo
                                {
                                    Range = tableRange,
                                    RowCount = rowCount,
                                    ColCount = colCount
                                };

                                string? tableName = element.Attribute("name")?.Value;
                                if (!string.IsNullOrEmpty(tableName))
                                {
                                    tables[tableName] = tableInfo;
                                }

                                lastTable = tableInfo;
                                currentRow = tableRange.Row + rowCount + 1;
                            }
                            else if (element.Name == "chart")
                            {
                                int chartWidth = int.TryParse(element.Attribute("chartWidth")?.Value, out int w) ? w : 10;
                                int chartHeight = int.TryParse(element.Attribute("chartHeight")?.Value, out int h) ? h : 15;

                                TableInfo chartTable = lastTable ?? new TableInfo { Range = sheet.Range["A1"], RowCount = 0, ColCount = 0 };

                                string? sourceTable = element.Attribute("sourceTable")?.Value;
                                if (!string.IsNullOrEmpty(sourceTable) && tables.ContainsKey(sourceTable))
                                {
                                    chartTable = tables[sourceTable];

                                    if (string.IsNullOrEmpty(element.Attribute("dataRange")?.Value))
                                    {
                                        string dataRange = $"{chartTable.Range.AddressLocal}:{sheet[chartTable.Range.Row + chartTable.RowCount - 1, chartTable.Range.Column + chartTable.ColCount - 1].AddressLocal}";
                                        element.SetAttributeValue("dataRange", dataRange);
                                    }
                                }

                                ChartComponent.AddChart(element, sheet, chartTable.Range, chartTable.RowCount, chartTable.ColCount, currentRow, chartWidth, chartHeight);

                                currentRow += chartHeight + 1;
                            }
                        
                    }

                        sheet.UsedRange.AutofitColumns();
                        sheet.UsedRange.AutofitRows();
                        sheet.Calculate();

                        pageCounter++;
                    }

                    using MemoryStream ms = new MemoryStream();
                    workbook.SaveAs(ms);
                    return ms.ToArray();
                }
                catch (Exception ex)
                {
                    throw new ExcelGenerationException("Excel oluşturulurken bir hata meydana geldi.", ex);
                }
            }
        }
    }
