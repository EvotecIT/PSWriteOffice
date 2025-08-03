using System.Diagnostics;
using System.IO;
using ClosedXML.Excel;

namespace PSWriteOffice.Services.Excel;

public static partial class ExcelDocumentService
{
    public static XLWorkbook LoadWorkbook(string filePath)
    {
        if (!File.Exists(filePath))
        {
            throw new FileNotFoundException($"File {filePath} doesn't exist.", filePath);
        }

        return new XLWorkbook(filePath);
    }

    public static XLWorkbook CreateWorkbook()
    {
        return new XLWorkbook();
    }

    public static void SaveWorkbook(XLWorkbook workbook, string filePath, bool show)
    {
        workbook.SaveAs(filePath);

        if (show)
        {
            var startInfo = new ProcessStartInfo
            {
                FileName = filePath,
                UseShellExecute = true
            };
            Process.Start(startInfo);
        }

        workbook.Dispose();
    }

    public static void CloseWorkbook(XLWorkbook workbook)
    {
        workbook.Dispose();
    }
}
