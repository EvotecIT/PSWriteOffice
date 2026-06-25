using OfficeIMO.CSV;
using PSWriteOffice.Services;

namespace PSWriteOffice.Cmdlets.Csv;

internal sealed class CsvPowerShellObjectProjector
{
    private string[]? _columns;
    private object?[]? _values;

    public void Reset()
    {
        _columns = null;
        _values = null;
    }

    public void WriteObject(object? value, CsvObjectWriter writer)
    {
        if (_columns != null &&
            _values != null &&
            PowerShellObjectNormalizer.TryProjectItemInto(value, _columns, _values))
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
}
