using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using OfficeIMO.Excel;
using PSWriteOffice.Services;

namespace PSWriteOffice.Services.Excel;

internal static class ExcelTabularInputService
{
    public static DataTable ToDataTable(
        IEnumerable<object?> input,
        string? tableName = null,
        bool copyExistingTables = true,
        PowerShellObjectNormalizerOptions? normalizerOptions = null)
    {
        if (input == null)
        {
            throw new ArgumentNullException(nameof(input));
        }

        normalizerOptions ??= PowerShellObjectNormalizerOptions.Default;

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
                return NormalizeTabularValues(dataTable, copyExistingTables, normalizerOptions);
            }

            if (single is DataView dataView)
            {
                return NormalizeTabularValues(dataView.ToTable(), copyExistingTable: false, normalizerOptions);
            }

            if (single is IDataReader reader)
            {
                var dataTableFromReader = string.IsNullOrWhiteSpace(tableName)
                    ? new DataTable()
                    : new DataTable(tableName);
                dataTableFromReader.Load(reader);
                return NormalizeTabularValues(dataTableFromReader, copyExistingTable: false, normalizerOptions);
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
                return NormalizeTabularValues(FromDataRows(rows), copyExistingTable: false, normalizerOptions);
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
                return NormalizeTabularValues(FromDataRows(rows), copyExistingTable: false, normalizerOptions);
            }
        }

        if (TryProjectToDataTable(items, tableName, normalizerOptions, out var projectedTable))
        {
            return projectedTable;
        }

        var normalized = PowerShellObjectNormalizer.NormalizeItems(items, normalizerOptions);
        return ExcelObjectDataTableBuilder.FromObjects(normalized, tableName ?? string.Empty);
    }

    /// <summary>Projects a normalized DataTable into dictionary rows accepted by generic report composers.</summary>
    public static object[] ToDictionaryRows(DataTable table)
    {
        if (table == null)
        {
            throw new ArgumentNullException(nameof(table));
        }

        var rows = new object[table.Rows.Count];
        for (var rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++)
        {
            var sourceRow = table.Rows[rowIndex];
            var row = new Dictionary<string, object?>(table.Columns.Count, StringComparer.OrdinalIgnoreCase);
            foreach (DataColumn column in table.Columns)
            {
                var value = sourceRow[column];
                row[column.ColumnName] = value == DBNull.Value ? null : value;
            }

            rows[rowIndex] = row;
        }

        return rows;
    }

    private static bool TryProjectToDataTable(
        IReadOnlyList<object?> items,
        string? tableName,
        PowerShellObjectNormalizerOptions? normalizerOptions,
        out DataTable table)
    {
        table = string.IsNullOrWhiteSpace(tableName)
            ? new DataTable()
            : new DataTable(tableName);

        string[]? columns = null;
        table.MinimumCapacity = Math.Max(table.MinimumCapacity, items.Count);
        table.BeginLoadData();
        try
        {
            foreach (var item in items)
            {
                if (!PowerShellObjectNormalizer.TryProjectItem(item, columns, out var projectedColumns, out var values, normalizerOptions))
                {
                    table.EndLoadData();
                    table = null!;
                    return false;
                }

                if (columns == null)
                {
                    columns = projectedColumns;
                    if (columns.Length == 0)
                    {
                        table.EndLoadData();
                        table = null!;
                        return false;
                    }

                    foreach (var column in columns)
                    {
                        table.Columns.Add(column, typeof(object));
                    }
                }

                for (var i = 0; i < values.Length; i++)
                {
                    values[i] ??= DBNull.Value;
                }

                table.Rows.Add(values);
            }
        }
        finally
        {
            table?.EndLoadData();
        }

        return columns != null;
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

    private static DataTable NormalizeTabularValues(
        DataTable source,
        bool copyExistingTable,
        PowerShellObjectNormalizerOptions options)
    {
        object?[][]? normalizedRows = null;
        for (var rowIndex = 0; rowIndex < source.Rows.Count; rowIndex++)
        {
            var sourceRow = source.Rows[rowIndex];
            object?[]? normalizedRow = null;
            for (var columnIndex = 0; columnIndex < source.Columns.Count; columnIndex++)
            {
                var value = sourceRow[columnIndex];
                var normalized = PowerShellObjectNormalizer.NormalizeCellValueForTable(value, options);
                if (normalizedRow == null && !CellValuesEqual(value, normalized))
                {
                    if (normalizedRows == null)
                    {
                        normalizedRows = new object?[source.Rows.Count][];
                        for (var previousRowIndex = 0; previousRowIndex < rowIndex; previousRowIndex++)
                        {
                            normalizedRows[previousRowIndex] = source.Rows[previousRowIndex].ItemArray.Cast<object?>().ToArray();
                        }
                    }
                    normalizedRow = new object?[source.Columns.Count];
                    for (var copyIndex = 0; copyIndex < columnIndex; copyIndex++)
                    {
                        normalizedRow[copyIndex] = sourceRow[copyIndex];
                    }
                }

                if (normalizedRow != null)
                {
                    normalizedRow[columnIndex] = normalized;
                }
            }

            if (normalizedRows != null)
            {
                normalizedRows[rowIndex] = normalizedRow ?? sourceRow.ItemArray.Cast<object?>().ToArray();
            }
        }

        if (normalizedRows == null)
        {
            return copyExistingTable ? source.Copy() : source;
        }

        var result = new DataTable(source.TableName)
        {
            CaseSensitive = source.CaseSensitive,
            Locale = source.Locale,
            Namespace = source.Namespace,
            Prefix = source.Prefix
        };
        foreach (DataColumn column in source.Columns)
        {
            result.Columns.Add(column.ColumnName, typeof(object));
        }

        foreach (var row in normalizedRows)
        {
            var values = row ?? Array.Empty<object?>();
            for (var index = 0; index < values.Length; index++)
            {
                values[index] ??= DBNull.Value;
            }
            result.Rows.Add(values);
        }

        return result;
    }

    private static bool CellValuesEqual(object? left, object? right)
    {
        if (ReferenceEquals(left, right))
        {
            return true;
        }

        return left?.GetType() == right?.GetType() && Equals(left, right);
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
