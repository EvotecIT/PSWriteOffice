using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.CSV;
using PSWriteOffice.Services;

namespace PSWriteOffice.Cmdlets.Csv;

internal sealed class CsvPowerShellObjectProjector
{
    private const int TypedProjectionRowThreshold = 1024;
    private string[]? _columns;
    private object?[]? _values;
    private string?[]? _textValues;
    private PowerShellObjectNormalizerOptions _normalizerOptions = PowerShellObjectNormalizerOptions.Default;
    private PowerShellObjectNormalizerOptions _objectNormalizerOptions = PowerShellObjectNormalizerOptions.Default;
    private bool _validateColumns;
    private KnownColumnProjectionMode _knownColumnProjectionMode;
    private int _knownColumnRowsWritten;
    private static readonly Func<ObjectProjectionState, int, object?> ObjectValueAccessor = static (state, index) => state.GetValue(index);

    public void Reset()
    {
        _columns = null;
        _values = null;
        _textValues = null;
        _validateColumns = false;
        _knownColumnProjectionMode = KnownColumnProjectionMode.Unknown;
        _knownColumnRowsWritten = 0;
    }

    public void UseCsvCulture(CultureInfo culture)
    {
        _normalizerOptions = new PowerShellObjectNormalizerOptions
        {
            Culture = culture ?? CultureInfo.InvariantCulture,
            FormatScalarValuesAsText = true
        };
        _objectNormalizerOptions = new PowerShellObjectNormalizerOptions
        {
            Culture = culture ?? CultureInfo.InvariantCulture,
            FormatScalarValuesAsText = false
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
        _knownColumnProjectionMode = KnownColumnProjectionMode.Unknown;
        _knownColumnRowsWritten = 0;
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
            if (_validateColumns)
            {
                ValidateFirstRowColumns(value, _columns);
            }

            if (PowerShellObjectNormalizer.TryPreparePSObjectTextProjection(value, out var ps))
            {
                if (_knownColumnProjectionMode == KnownColumnProjectionMode.Unknown)
                {
                    _knownColumnProjectionMode = ShouldUseTypedProjection(_columns)
                        ? KnownColumnProjectionMode.Typed
                        : KnownColumnProjectionMode.Text;
                }

                if (_knownColumnProjectionMode == KnownColumnProjectionMode.Typed &&
                    _knownColumnRowsWritten >= TypedProjectionRowThreshold)
                {
                    WriteProjectedRow(writer, _columns, new ObjectProjectionState(ps!, _columns, _objectNormalizerOptions));
                    _knownColumnRowsWritten++;
                    return;
                }

                if (_textValues != null &&
                    PowerShellObjectNormalizer.TryProjectPSObjectTextIntoKnownColumns(value, _columns, _textValues, _normalizerOptions))
                {
                    WriteProjectedTextRow(writer, _columns, _textValues);
                    _knownColumnRowsWritten++;
                    return;
                }
            }

            if (PowerShellObjectNormalizer.TryProjectPSObjectIntoKnownColumns(value, _columns, _values, _normalizerOptions))
            {
                WriteProjectedRow(writer, _columns, _values);
                _knownColumnRowsWritten++;
                return;
            }

            if (!TryProjectIntoExistingColumns(value, _columns, _values))
            {
                Array.Clear(_values, 0, _values.Length);
            }

            WriteProjectedRow(writer, _columns, _values);
            _knownColumnRowsWritten++;
            return;
        }

        if (PowerShellObjectNormalizer.TryProjectItem(value, null, out var columns, out var values, _normalizerOptions))
        {
            _columns = columns;
            _values = new object?[columns.Length];
            _textValues = new string?[columns.Length];
            if (TryCopyTextValues(values, _textValues))
            {
                writer.WriteTextRow(_columns, _textValues);
                return;
            }

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

        writer.WriteTextRow(columns, values);
    }

    private static bool TryCopyTextValues(object?[] values, string?[] textValues)
    {
        for (var i = 0; i < values.Length; i++)
        {
            if (values[i] is string text)
            {
                textValues[i] = text;
                continue;
            }

            if (values[i] == null)
            {
                textValues[i] = null;
                continue;
            }

            Array.Clear(textValues, 0, textValues.Length);
            return false;
        }

        return true;
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

    private static void WriteProjectedRow(CsvObjectWriter writer, string[] columns, ObjectProjectionState state)
    {
        if (writer.HasRows)
        {
            writer.WriteTrustedRow(columns.Length, state, ObjectValueAccessor);
            return;
        }

        writer.WriteRow(columns, columns.Length, state, ObjectValueAccessor);
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

    private static bool ShouldUseTypedProjection(string[] columns)
    {
        return columns.Length > 0;
    }

    private enum KnownColumnProjectionMode
    {
        Unknown,
        Text,
        Typed
    }

    private readonly struct ObjectProjectionState
    {
        private readonly PSObject _ps;
        private readonly string[] _columns;
        private readonly PowerShellObjectNormalizerOptions _options;

        public ObjectProjectionState(PSObject ps, string[] columns, PowerShellObjectNormalizerOptions options)
        {
            _ps = ps;
            _columns = columns;
            _options = options;
        }

        public object? GetValue(int index)
        {
            return PowerShellObjectNormalizer.ProjectPSObjectValue(_ps, _columns[index], _options);
        }
    }
}
