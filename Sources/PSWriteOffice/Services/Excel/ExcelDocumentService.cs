using System;
using System.Collections.Generic;

namespace PSWriteOffice.Services.Excel;

public static partial class ExcelDocumentService
{
    public static object ConvertWorksheetData(IEnumerable<IDictionary<string, object?>> rows, Type? type, bool asDataTable)
        => asDataTable
            ? BuildDataTable(rows)
            : type != null
                ? MapRowsToType(rows, type)
                : rows;
}
