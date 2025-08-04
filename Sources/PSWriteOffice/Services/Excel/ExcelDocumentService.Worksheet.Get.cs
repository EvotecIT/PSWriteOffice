using System.Collections.Generic;
using ClosedXML.Excel;

namespace PSWriteOffice.Services.Excel;

public static partial class ExcelDocumentService
{
    public static IXLWorksheet? GetWorksheet(XLWorkbook workbook, string worksheetName)
    {
        return workbook.Worksheets.TryGetWorksheet(worksheetName, out var worksheet) ? worksheet : null;
    }

    public static IXLWorksheet? GetWorksheet(XLWorkbook workbook, int index)
    {
        return index >= 1 && index <= workbook.Worksheets.Count ? workbook.Worksheet(index) : null;
    }

    public static IEnumerable<IXLWorksheet> GetWorksheets(XLWorkbook workbook)
    {
        return workbook.Worksheets;
    }
}
