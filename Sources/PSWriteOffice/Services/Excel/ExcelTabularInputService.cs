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
            ProjectObject = item => ProjectPowerShellObject(item, normalizerOptions)
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

    private static IReadOnlyDictionary<string, object?>? ProjectPowerShellObject(object? item, PowerShellObjectNormalizerOptions? normalizerOptions)
    {
        if (!PowerShellObjectNormalizer.TryProjectItem(item, null, out var columns, out var values, normalizerOptions) ||
            columns.Length == 0)
        {
            return null;
        }

        var row = new Dictionary<string, object?>(columns.Length, StringComparer.OrdinalIgnoreCase);
        for (var i = 0; i < columns.Length; i++)
        {
            row[columns[i]] = values[i];
        }

        return row;
    }
}
