using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using OfficeIMO.Data;
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

        var table = TabularDataTableBuilder.FromItems(input, new TabularDataOptions
        {
            TableName = tableName,
            CopyExistingDataTable = copyExistingTables,
            ColumnDiscoveryMode = TabularColumnDiscoveryMode.FirstRow,
            UnwrapValue = Unwrap,
            ProjectObject = (item, columns) => ProjectPowerShellObject(item, columns, normalizerOptions)
        });

        if (table.Columns.Count == 0 && table.Rows.Count == 0)
        {
            throw new ArgumentException("Provide at least one data row.", nameof(input));
        }

        return table;
    }

    public static DataSet? TryGetSingleDataSet(IEnumerable<object?> input)
    {
        if (input == null)
        {
            throw new ArgumentNullException(nameof(input));
        }

        return TabularDataTableBuilder.TryGetSingleDataSet(input, Unwrap);
    }

    public static IDataReader? TryGetSingleDataReader(IEnumerable<object?> input)
    {
        if (input == null)
        {
            throw new ArgumentNullException(nameof(input));
        }

        return TabularDataTableBuilder.TryGetSingleDataReader(input, Unwrap);
    }

    private static object? Unwrap(object? item)
    {
        if (item is System.Management.Automation.PSObject psObject)
        {
            return psObject.BaseObject;
        }

        return item;
    }

    private static IReadOnlyDictionary<string, object?>? ProjectPowerShellObject(
        object? item,
        IReadOnlyList<string>? columns,
        PowerShellObjectNormalizerOptions? normalizerOptions)
    {
        var knownColumns = columns as string[] ?? columns?.ToArray();
        if (!PowerShellObjectNormalizer.TryProjectItem(item, knownColumns, out var projectedColumns, out var values, normalizerOptions) ||
            projectedColumns.Length == 0)
        {
            return null;
        }

        var row = new Dictionary<string, object?>(projectedColumns.Length, StringComparer.OrdinalIgnoreCase);
        for (var i = 0; i < projectedColumns.Length; i++)
        {
            row[projectedColumns[i]] = values[i];
        }

        return row;
    }
}
