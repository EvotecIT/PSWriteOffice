using System.Collections.Generic;
using ClosedXML.Excel;

namespace PSWriteOffice.Services.Excel;

public static partial class ExcelDocumentService
{
    public static void ApplyFormulas(IXLWorksheet worksheet, IDictionary<string, string> formulas)
    {
        foreach (var kvp in formulas)
        {
            worksheet.Cell(kvp.Key).FormulaA1 = kvp.Value;
        }
    }
}
