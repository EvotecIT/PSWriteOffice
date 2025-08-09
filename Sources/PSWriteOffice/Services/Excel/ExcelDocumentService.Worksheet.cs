using System;
using ClosedXML.Excel;

namespace PSWriteOffice.Services.Excel;

public enum WorksheetExistOption
{
    Replace,
    Skip,
    Rename
}

public static partial class ExcelDocumentService
{
    public static IXLWorksheet AddWorksheet(XLWorkbook workbook, string worksheetName, WorksheetExistOption option, XLColor? tabColor = null)
    {
        if (workbook.Worksheets.Contains(worksheetName))
        {
            switch (option)
            {
                case WorksheetExistOption.Replace:
                    workbook.Worksheet(worksheetName).Delete();
                    break;
                case WorksheetExistOption.Rename:
                    worksheetName = $"Sheet{Guid.NewGuid():N}".Substring(0, 8);
                    break;
                case WorksheetExistOption.Skip:
                    var existing = workbook.Worksheet(worksheetName);
                    if (tabColor != null)
                    {
                        existing.TabColor = tabColor;
                    }
                    return existing;
            }
        }

        var worksheet = workbook.Worksheets.Add(worksheetName);
        if (tabColor != null)
        {
            worksheet.TabColor = tabColor;
        }
        return worksheet;
    }

    public static void AutoSizeColumns(IXLWorksheet worksheet)
    {
        worksheet.Columns().AdjustToContents();
    }

    public static void FreezeTopRow(IXLWorksheet worksheet)
    {
        worksheet.SheetView.FreezeRows(1);
    }

    public static void FreezeFirstColumn(IXLWorksheet worksheet)
    {
        worksheet.SheetView.FreezeColumns(1);
    }
}
