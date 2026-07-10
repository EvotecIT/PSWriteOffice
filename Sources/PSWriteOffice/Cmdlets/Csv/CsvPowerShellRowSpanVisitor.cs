#nullable enable

#if NET8_0_OR_GREATER
using System;
using System.Collections.Generic;
using System.Management.Automation;
using OfficeIMO.CSV;

namespace PSWriteOffice.Cmdlets.Csv;

/// <summary>Builds PowerShell row objects directly from OfficeIMO CSV field spans.</summary>
internal struct CsvPowerShellRowSpanVisitor : ICsvRowFieldSpanVisitor
{
    private readonly CsvPowerShellRowWriter _rowWriter;
    private readonly PSCmdlet _cmdlet;
    private string[]? _header;
    private PSObject? _current;
    private int _propertyCount;
    private bool _prevalidated;

    internal CsvPowerShellRowSpanVisitor(CsvPowerShellRowWriter rowWriter, PSCmdlet cmdlet)
    {
        _rowWriter = rowWriter;
        _cmdlet = cmdlet;
        _header = null;
        _current = null;
        _propertyCount = 0;
        _prevalidated = false;
    }

    public void BeginRow(IReadOnlyList<string> header, int rowIndex)
    {
        _header ??= _rowWriter.GetOutputHeader(header);
        _prevalidated = _rowWriter.PrevalidatedOutputProperties;
        _current = CsvPowerShellRowWriter.PowerShellObjectFactory.Create(_header.Length);
        _propertyCount = 0;
    }

    public void VisitField(int rowIndex, int fieldIndex, ReadOnlySpan<char> value)
    {
        AddField(fieldIndex, value.ToString());
    }

    public void VisitFieldValue(int rowIndex, int fieldIndex, string value)
    {
        AddField(fieldIndex, value);
    }

    public void EndRow(int rowIndex, int fieldCount)
    {
        var current = _current ?? throw new InvalidOperationException("CSV row was not started.");
        var header = _header ?? throw new InvalidOperationException("CSV header was not initialized.");
        if (_propertyCount < header.Length)
        {
            CsvPowerShellRowWriter.AddMissingNoteProperties(
                current,
                header,
                _propertyCount,
                header.Length,
                _prevalidated);
        }

        _cmdlet.WriteObject(current);
        _current = null;
    }

    private void AddField(int fieldIndex, string value)
    {
        var current = _current ?? throw new InvalidOperationException("CSV row was not started.");
        var header = _header ?? throw new InvalidOperationException("CSV header was not initialized.");
        if ((uint)fieldIndex >= (uint)header.Length)
        {
            return;
        }

        current.Properties.Add(
            new PSNoteProperty(header[fieldIndex], value),
            _prevalidated);
        _propertyCount++;
    }
}
#endif
