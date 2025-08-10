using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using ClosedXML.Excel;

namespace PSWriteOffice.Services.Excel;

public static partial class ExcelDocumentService
{
    public static IXLTable InsertTable(IXLWorksheet worksheet, IEnumerable<IDictionary<string, object?>> data, int row, int column,
        XLTableTheme theme, bool showRowStripes, bool showColumnStripes, bool showAutoFilter, bool showHeaderRow, bool showTotalsRow,
        bool emphasizeFirstColumn, bool emphasizeLastColumn, XLTransposeOptions? transpose)
    {
        // Convert the dictionary data to a DataTable that ClosedXML can handle properly
        var dataTable = new System.Data.DataTable();
        var dataList = data.ToList();
        
        if (dataList.Count > 0)
        {
            // Add columns based on the first row
            foreach (var key in dataList[0].Keys)
            {
                dataTable.Columns.Add(key, typeof(object));
            }
            
            // Add rows
            foreach (var rowData in dataList)
            {
                var dataRow = dataTable.NewRow();
                foreach (var kvp in rowData)
                {
                    if (dataTable.Columns.Contains(kvp.Key))
                    {
                        dataRow[kvp.Key] = kvp.Value ?? DBNull.Value;
                    }
                }
                dataTable.Rows.Add(dataRow);
            }
        }
        
        var table = worksheet.Cell(row, column).InsertTable(dataTable);

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
