#if NET8_0_OR_GREATER
using System;
using System.Collections.Generic;
using System.Management.Automation;
using OfficeIMO.CSV;

namespace PSWriteOffice.Cmdlets.Csv;

internal struct CsvPowerShellRowFieldSpanVisitor : ICsvRowFieldSpanVisitor
{
    private readonly CsvPowerShellRowWriter _rowWriter;
    private readonly PSCmdlet _cmdlet;
    private readonly bool _asHashtable;

    public CsvPowerShellRowFieldSpanVisitor(CsvPowerShellRowWriter rowWriter, PSCmdlet cmdlet, bool asHashtable)
    {
        _rowWriter = rowWriter;
        _cmdlet = cmdlet;
        _asHashtable = asHashtable;
    }

    public readonly void BeginRow(IReadOnlyList<string> header, int rowIndex)
    {
        _rowWriter.BeginSpanRow(header, _asHashtable);
    }

    public readonly void VisitField(int rowIndex, int fieldIndex, ReadOnlySpan<char> value)
    {
        _rowWriter.WriteSpanField(fieldIndex, value);
    }

    public readonly void VisitFieldValue(int rowIndex, int fieldIndex, string value)
    {
        _rowWriter.WriteSpanFieldValue(fieldIndex, value);
    }

    public readonly void EndRow(int rowIndex, int fieldCount)
    {
        _rowWriter.EndSpanRow(_cmdlet);
    }
}
#endif
