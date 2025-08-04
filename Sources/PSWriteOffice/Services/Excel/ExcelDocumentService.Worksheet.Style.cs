using ClosedXML.Excel;

namespace PSWriteOffice.Services.Excel;

public static partial class ExcelDocumentService
{
    public static void SetWorksheetTabColor(IXLWorksheet worksheet, XLColor color)
    {
        worksheet.SetTabColor(color);
    }
}
