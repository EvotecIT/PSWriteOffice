using System.Collections.Generic;
using ClosedXML.Excel;

namespace PSWriteOffice.Services.Excel;

public static partial class ExcelDocumentService
{
    public static IXLPivotTable AddPivotTable(
        IXLWorksheet worksheet,
        string name,
        string sourceRange,
        string targetCell,
        IEnumerable<string>? rowLabels,
        IEnumerable<string>? columnLabels,
        IDictionary<string, XLPivotSummary>? values)
    {
        var pivot = worksheet.PivotTables.Add(name, worksheet.Cell(targetCell), worksheet.Range(sourceRange));

        if (rowLabels != null)
        {
            foreach (var label in rowLabels)
            {
                pivot.RowLabels.Add(label);
            }
        }

        if (columnLabels != null)
        {
            foreach (var label in columnLabels)
            {
                pivot.ColumnLabels.Add(label);
            }
        }

        if (values != null)
        {
            foreach (var kvp in values)
            {
                pivot.Values.Add(kvp.Key).SetSummaryFormula(kvp.Value);
            }
        }

        return pivot;
    }
}
