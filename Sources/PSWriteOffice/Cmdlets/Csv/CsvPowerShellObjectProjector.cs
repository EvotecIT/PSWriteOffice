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
    private bool _validateColumns;

    public void Reset()
    {
        _columns = null;
        _values = null;
        _validateColumns = false;
    }

    public void UseColumns(IReadOnlyList<string> columns, bool validateColumns)
    {
        if (columns == null)
        {
            throw new ArgumentNullException(nameof(columns));
        }

        _columns = columns.ToArray();
        _values = new object?[_columns.Length];
        _validateColumns = validateColumns;
    }

    public void ValidateObjectColumns(object? value, IReadOnlyList<string> columns)
    {
        if (columns == null)
        {
            throw new ArgumentNullException(nameof(columns));
        }

        ValidateFirstRowColumns(value, columns);
    }

    public bool CanProjectColumns(object? value, IReadOnlyList<string> columns)
    {
        if (columns == null)
        {
            throw new ArgumentNullException(nameof(columns));
        }

        return TryGetProjectableColumns(value, columns, out _);
    }

    public void WriteObject(object? value, CsvObjectWriter writer)
    {
        if (_columns != null &&
            _values != null)
        {
            if (!TryProjectIntoExistingColumns(value, _columns, _values))
            {
                Array.Clear(_values, 0, _values.Length);
            }

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
        if (_validateColumns)
        {
            ValidateFirstRowColumns(value, columns);
        }

        return PowerShellObjectNormalizer.TryProjectItemInto(value, columns, values);
    }

    private static void ValidateFirstRowColumns(object? value, IReadOnlyList<string> columns)
    {
        if (TryGetProjectableColumns(value, columns, out var missingColumn))
        {
            return;
        }

        if (missingColumn != null)
        {
            throw new CsvException($"Cannot append CSV because the input object is missing the existing column '{missingColumn}'. Use -Force to append with blank values for missing columns.");
        }
    }

    private static bool TryGetProjectableColumns(object? value, IReadOnlyList<string> columns, out string? missingColumn)
    {
        missingColumn = null;
        if (!PowerShellObjectNormalizer.TryProjectItem(value, null, out var sourceColumns, out _))
        {
            return false;
        }

        var sourceColumnSet = new HashSet<string>(sourceColumns, StringComparer.OrdinalIgnoreCase);
        foreach (var column in columns)
        {
            if (!sourceColumnSet.Contains(column))
            {
                missingColumn = column;
                return false;
            }
        }

        return true;
    }
}
