using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using OfficeIMO.Excel;
using PSWriteOffice.Services;

namespace PSWriteOffice.Services.Excel;

internal static class ExcelTabularInputService
{
    public static DataTable ToDataTable(IEnumerable<object?> input, string? tableName = null, bool copyExistingTables = true)
    {
        if (input == null)
        {
            throw new ArgumentNullException(nameof(input));
        }

        var items = new List<object?>();
        foreach (var item in input)
        {
            if (item == null)
            {
                continue;
            }

            items.Add(item);
        }

        if (items.Count == 0)
        {
            throw new ArgumentException("Provide at least one data row.", nameof(input));
        }

        if (items.Count == 1)
        {
            var single = Unwrap(items[0]);
            if (single is DataTable dataTable)
            {
                return copyExistingTables ? dataTable.Copy() : dataTable;
            }

            if (single is DataView dataView)
            {
                return dataView.ToTable();
            }

            if (single is IDataReader reader)
            {
                var dataTableFromReader = string.IsNullOrWhiteSpace(tableName)
                    ? new DataTable()
                    : new DataTable(tableName);
                dataTableFromReader.Load(reader);
                return dataTableFromReader;
            }
        }

        var first = Unwrap(items[0]);
        if (first is DataRow firstRow)
        {
            var rows = new List<DataRow>(items.Count) { firstRow };
            for (var i = 1; i < items.Count; i++)
            {
                if (Unwrap(items[i]) is not DataRow row)
                {
                    rows.Clear();
                    break;
                }

                rows.Add(row);
            }

            if (rows.Count > 0)
            {
                return FromDataRows(rows);
            }
        }

        if (first is DataRowView firstRowView)
        {
            var rows = new List<DataRow>(items.Count) { firstRowView.Row };
            for (var i = 1; i < items.Count; i++)
            {
                if (Unwrap(items[i]) is not DataRowView rowView)
                {
                    rows.Clear();
                    break;
                }

                rows.Add(rowView.Row);
            }

            if (rows.Count > 0)
            {
                return FromDataRows(rows);
            }
        }

        var normalized = PowerShellObjectNormalizer.NormalizeItems(items);
        return ObjectDataTableBuilder.FromObjects(normalized, tableName ?? string.Empty);
    }

    public static DataSet? TryGetSingleDataSet(IEnumerable<object?> input)
    {
        if (input == null)
        {
            throw new ArgumentNullException(nameof(input));
        }

        DataSet? dataSet = null;
        var count = 0;
        foreach (var item in input)
        {
            if (item == null)
            {
                continue;
            }

            count++;
            if (count > 1)
            {
                return null;
            }

            dataSet = Unwrap(item) as DataSet;
        }

        return count == 1 ? dataSet : null;
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
