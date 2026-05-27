using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Management.Automation;
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

        if (TryBuildFromPowerShellObjects(items, tableName, out var powerShellObjectTable))
        {
            return powerShellObjectTable;
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

    public static IDataReader? TryGetSingleDataReader(IEnumerable<object?> input)
    {
        if (input == null)
        {
            throw new ArgumentNullException(nameof(input));
        }

        IDataReader? reader = null;
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

            reader = Unwrap(item) as IDataReader;
        }

        return count == 1 ? reader : null;
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

    private static bool TryBuildFromPowerShellObjects(IReadOnlyList<object?> items, string? tableName, out DataTable table)
    {
        table = new DataTable();
        var result = string.IsNullOrWhiteSpace(tableName)
            ? new DataTable()
            : new DataTable(tableName);

        if (items.Count == 0)
        {
            return false;
        }

        if (items[0] is not PSObject firstObject)
        {
            return false;
        }

        var columns = GetPowerShellPropertyNames(firstObject);
        if (columns.Count == 0)
        {
            return false;
        }

        foreach (var column in columns)
        {
            result.Columns.Add(column, typeof(object));
        }

        var columnIndexes = CreateColumnIndexMap(columns);

        var values = new object?[columns.Count];
        result.BeginLoadData();
        try
        {
            foreach (var item in items)
            {
                if (item is not PSObject psObject)
                {
                    return false;
                }

                if (!TryReadPowerShellObjectRow(psObject, columns.Count, columnIndexes, values))
                {
                    return false;
                }

                result.Rows.Add(values);
            }
        }
        finally
        {
            result.EndLoadData();
        }

        table = result;
        return true;
    }

    private static List<string> GetPowerShellPropertyNames(PSObject psObject)
    {
        var names = new List<string>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var property in psObject.Properties)
        {
            if (!IsExportablePowerShellProperty(property))
            {
                continue;
            }

            var name = property.Name;
            if (string.IsNullOrWhiteSpace(name) || !seen.Add(name))
            {
                continue;
            }

            names.Add(name);
        }

        return names;
    }

    private static bool TryReadPowerShellObjectRow(PSObject psObject, int columnCount, IReadOnlyDictionary<string, int> columnIndexes, object?[] values)
    {
        Array.Clear(values, 0, values.Length);
        var matched = 0;
        foreach (var property in psObject.Properties)
        {
            if (!IsExportablePowerShellProperty(property))
            {
                continue;
            }

            if (!columnIndexes.TryGetValue(property.Name, out var columnIndex))
            {
                return false;
            }

            values[columnIndex] = property.Value ?? DBNull.Value;
            matched++;
        }

        return matched == columnCount;
    }

    private static Dictionary<string, int> CreateColumnIndexMap(IReadOnlyList<string> columns)
    {
        var columnIndexes = new Dictionary<string, int>(columns.Count, StringComparer.OrdinalIgnoreCase);
        for (var columnIndex = 0; columnIndex < columns.Count; columnIndex++)
        {
            columnIndexes[columns[columnIndex]] = columnIndex;
        }

        return columnIndexes;
    }

    private static bool IsExportablePowerShellProperty(PSPropertyInfo property)
    {
        return property.MemberType is PSMemberTypes.NoteProperty or PSMemberTypes.Property;
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
