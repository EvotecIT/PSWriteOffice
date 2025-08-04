using ClosedXML.Excel;

namespace PSWriteOffice.Services.Excel;

public static partial class ExcelDocumentService
{
    public static IXLCell GetCell(IXLWorksheet worksheet, int row, int column)
    {
        return worksheet.Cell(row, column);
    }
}
