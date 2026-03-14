using System;
using OfficeIMO.Excel;

namespace PSWriteOffice.Services.Excel;

internal static class ExcelSheetResolver
{
    public static ExcelSheet Resolve(ExcelDocument document, string? sheetName, int? sheetIndex)
    {
        if (document == null)
        {
            throw new ArgumentNullException(nameof(document));
        }

        if (!string.IsNullOrWhiteSpace(sheetName))
        {
            return document[sheetName!];
        }

        if (sheetIndex.HasValue)
        {
            if (sheetIndex.Value < 0 || sheetIndex.Value >= document.Sheets.Count)
            {
                throw new ArgumentOutOfRangeException(nameof(sheetIndex), "SheetIndex is out of range.");
            }

            return document.Sheets[sheetIndex.Value];
        }

        if (document.Sheets.Count == 0)
        {
            throw new InvalidOperationException("Workbook contains no worksheets.");
        }

        return document.Sheets[0];
    }
}
