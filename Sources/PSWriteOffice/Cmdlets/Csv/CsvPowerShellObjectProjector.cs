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
    private string?[]? _textValues;
    private PowerShellObjectNormalizerOptions _normalizerOptions = PowerShellObjectNormalizerOptions.Default;
    private bool _allowTrustedTextRows = true;
    private bool _validateColumns;

    public void Reset()
    {
        _columns = null;
        _values = null;
        _textValues = null;
        _validateColumns = false;
        _allowTrustedTextRows = true;
    }

    public void UseCsvOptions(CsvSaveOptions options)
    {
        if (options == null)
        {
            throw new ArgumentNullException(nameof(options));
        }

        _allowTrustedTextRows = options.DateTimeFormat == null && !options.UseUtc;
        _normalizerOptions = new PowerShellObjectNormalizerOptions
        {
            Culture = options.Culture,
            FormatScalarValuesAsText = _allowTrustedTextRows
        };
    }

    public void UseColumns(IReadOnlyList<string> columns, bool validateColumns)
    {
        if (columns == null)
        {
            throw new ArgumentNullException(nameof(columns));
        }

        _columns = columns.ToArray();
        _values = new object?[_columns.Length];
        _textValues = new string?[_columns.Length];
        _validateColumns = validateColumns;
    }

    public IReadOnlyList<string>? CurrentColumns => _columns;

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
            if (_validateColumns)
            {
                ValidateFirstRowColumns(value, _columns);
            }

            if (_allowTrustedTextRows &&
                _textValues != null &&
                PowerShellObjectNormalizer.TryProjectPSObjectTextIntoKnownColumns(value, _columns, _textValues, _normalizerOptions))
            {
                WriteProjectedTextRow(writer, _columns, _textValues);
                return;
            }

            if (PowerShellObjectNormalizer.TryProjectPSObjectIntoKnownColumns(value, _columns, _values, _normalizerOptions))
            {
                WriteProjectedRow(writer, _columns, _values);
                return;
            }

            if (!TryProjectIntoExistingColumns(value, _columns, _values))
            {
                Array.Clear(_values, 0, _values.Length);
            }

            WriteProjectedRow(writer, _columns, _values);
            return;
        }

        if (PowerShellObjectNormalizer.TryProjectItem(value, null, out var columns, out var values, _normalizerOptions))
        {
            _columns = columns;
            _values = new object?[columns.Length];
            _textValues = new string?[columns.Length];
            writer.WriteRow(_columns, values);
            return;
        }

        writer.WriteObject(PowerShellObjectNormalizer.NormalizeItem(value));
    }

    private static void WriteProjectedTextRow(CsvObjectWriter writer, string[] columns, string?[] values)
    {
        if (writer.HasRows)
        {
            writer.WriteTrustedTextRow(values);
            return;
        }

        writer.WriteRow(columns, values);
    }

    private static void WriteProjectedRow(CsvObjectWriter writer, string[] columns, object?[] values)
    {
        if (writer.HasRows)
        {
            writer.WriteTrustedRow(values);
            return;
        }

        writer.WriteRow(columns, values);
    }

    private bool TryProjectIntoExistingColumns(object? value, string[] columns, object?[] values)
    {
        if (_validateColumns)
        {
            ValidateFirstRowColumns(value, columns);
        }

        return PowerShellObjectNormalizer.TryProjectItemInto(value, columns, values, _normalizerOptions);
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
        if (!PowerShellObjectNormalizer.TryProjectItem(value, null, out var sourceColumns, out _, PowerShellObjectNormalizerOptions.Default))
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
