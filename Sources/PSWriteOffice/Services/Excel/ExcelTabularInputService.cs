using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using OfficeIMO.Excel;
using PSWriteOffice.Services;

namespace PSWriteOffice.Services.Excel;

internal static class ExcelTabularInputService
{
    public static DataTable ToDataTable(IEnumerable<object?> input, string tableName = "Data")
    {
        if (input == null)
        {
            throw new ArgumentNullException(nameof(input));
        }

        var items = input.Where(item => item != null).ToList();
        if (items.Count == 0)
        {
            throw new ArgumentException("Provide at least one data row.", nameof(input));
        }

        if (items.Count == 1)
        {
            var single = Unwrap(items[0]);
            if (single is DataTable dataTable)
            {
                return dataTable.Copy();
            }

            if (single is DataView dataView)
            {
                return dataView.ToTable();
            }

            if (single is IDataReader reader)
            {
                var dataTableFromReader = new DataTable(tableName);
                dataTableFromReader.Load(reader);
                return dataTableFromReader;
            }
        }

        if (items.All(item => Unwrap(item) is DataRow))
        {
            return FromDataRows(items.Select(item => (DataRow)Unwrap(item)!).ToList());
        }

        if (items.All(item => Unwrap(item) is DataRowView))
        {
            return FromDataRows(items.Select(item => ((DataRowView)Unwrap(item)!).Row).ToList());
        }

        var normalized = PowerShellObjectNormalizer.NormalizeItems(items);
        return ObjectDataTableBuilder.FromObjects(normalized, tableName);
    }

    public static DataSet? TryGetSingleDataSet(IEnumerable<object?> input)
    {
        if (input == null)
        {
            throw new ArgumentNullException(nameof(input));
        }

        var items = input.Where(item => item != null).Select(Unwrap).ToList();
        return items.Count == 1 ? items[0] as DataSet : null;
    }

    private static DataTable FromDataRows(IReadOnlyList<DataRow> rows)
    {
        if (rows.Count == 0)
        {
            throw new ArgumentException("Provide at least one data row.", nameof(rows));
        }

        var source = rows[0].Table;
        var result = source.Clone();
        foreach (var row in rows)
        {
            if (!ReferenceEquals(row.Table, source))
            {
                throw new InvalidOperationException("DataRow inputs must come from the same DataTable.");
            }

            result.ImportRow(row);
        }

        return result;
    }

    private static object? Unwrap(object? item)
    {
        if (item is System.Management.Automation.PSObject psObject)
        {
            return psObject.BaseObject;
        }

        return item;
    }
}
