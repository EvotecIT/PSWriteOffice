using System.Collections.Generic;
using ClosedXML.Excel;

namespace PSWriteOffice.Services.Excel;

public static partial class ExcelDocumentService
{
    public static IXLTable InsertTable(IXLWorksheet worksheet, IEnumerable<IDictionary<string, object?>> data, int row, int column,
        XLTableTheme theme, bool showRowStripes, bool showColumnStripes, bool showAutoFilter, bool showHeaderRow, bool showTotalsRow,
        bool emphasizeFirstColumn, bool emphasizeLastColumn, XLTransposeOptions? transpose)
    {
        var table = worksheet.Cell(row, column).InsertTable(data);

        if (transpose.HasValue)
        {
            table.Transpose(transpose.Value);
        }

        table.ShowRowStripes = showRowStripes;
        table.ShowColumnStripes = showColumnStripes;
        table.ShowAutoFilter = showAutoFilter;
        table.ShowHeaderRow = showHeaderRow;
        table.ShowTotalsRow = showTotalsRow;
        table.EmphasizeFirstColumn = emphasizeFirstColumn;
        table.EmphasizeLastColumn = emphasizeLastColumn;

        if (theme != XLTableTheme.None)
        {
            table.Theme = theme;
        }

        return table;
    }
}
