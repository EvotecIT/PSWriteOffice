using System;
using System.Collections.Generic;
using System.Linq;
using OfficeIMO.CSV;
using PSWriteOffice.Services;

namespace PSWriteOffice.Cmdlets.Csv;

internal sealed class CsvPowerShellObjectProjector
{
    private string[]? _columns;
    private object?[]? _values;
    private bool _validateFirstRowColumns;

    public void Reset()
    {
        _columns = null;
        _values = null;
        _validateFirstRowColumns = false;
    }

    public void UseColumns(IReadOnlyList<string> columns, bool validateFirstRowColumns)
    {
        if (columns == null)
        {
            throw new ArgumentNullException(nameof(columns));
        }

        _columns = columns.ToArray();
        _values = new object?[_columns.Length];
        _validateFirstRowColumns = validateFirstRowColumns;
    }

    public void WriteObject(object? value, CsvObjectWriter writer)
    {
        if (_columns != null &&
            _values != null &&
            TryProjectIntoExistingColumns(value, _columns, _values))
        {
            writer.WriteRow(_columns, _values);
            return;
        }

        if (PowerShellObjectNormalizer.TryProjectItem(value, null, out var columns, out var values))
        {
            _columns = columns;
            _values = new object?[columns.Length];
            writer.WriteRow(_columns, values);
            return;
        }

        writer.WriteObject(PowerShellObjectNormalizer.NormalizeItem(value));
    }

    private bool TryProjectIntoExistingColumns(object? value, string[] columns, object?[] values)
    {
        if (_validateFirstRowColumns)
        {
            ValidateFirstRowColumns(value, columns);
            _validateFirstRowColumns = false;
        }

        return PowerShellObjectNormalizer.TryProjectItemInto(value, columns, values);
    }

    private static void ValidateFirstRowColumns(object? value, string[] columns)
    {
        if (!PowerShellObjectNormalizer.TryProjectItem(value, null, out var sourceColumns, out _))
        {
            return;
        }

        var sourceColumnSet = new HashSet<string>(sourceColumns, StringComparer.OrdinalIgnoreCase);
        foreach (var column in columns)
        {
            if (!sourceColumnSet.Contains(column))
            {
                throw new CsvException($"Cannot append CSV because the input object is missing the existing column '{column}'. Use -Force to append with blank values for missing columns.");
            }
        }
    }
}
