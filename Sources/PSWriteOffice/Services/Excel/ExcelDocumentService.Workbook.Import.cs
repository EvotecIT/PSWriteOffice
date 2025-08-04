using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;

namespace PSWriteOffice.Services.Excel;

public static partial class ExcelDocumentService
{
    public static IDictionary<string, IList<IDictionary<string, object?>>> ImportWorkbook(string filePath, IEnumerable<string>? worksheetNames = null)
    {
        using var workbook = LoadWorkbook(filePath);
        var result = new Dictionary<string, IList<IDictionary<string, object?>>>();

        foreach (var worksheet in workbook.Worksheets)
        {
            if (worksheetNames != null && worksheetNames.Any() && !worksheetNames.Contains(worksheet.Name))
            {
                continue;
            }

            var rows = GetWorksheetData(worksheet).ToList();
            result[worksheet.Name] = rows;
        }

        return result;
    }
}
