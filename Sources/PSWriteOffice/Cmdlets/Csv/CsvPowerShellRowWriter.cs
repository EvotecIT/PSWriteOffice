using System;
using System.Collections.Generic;
using System.Management.Automation;
using OfficeIMO.CSV;

namespace PSWriteOffice.Cmdlets.Csv;

internal sealed class CsvPowerShellRowWriter
{
    private bool _prevalidatedOutputProperties;
    private string[]? _outputHeader;

    public void Reset()
    {
        _outputHeader = null;
        _prevalidatedOutputProperties = false;
    }

    public void WriteDocumentRows(CsvDocument document, bool asHashtable, Action<object?> writeObject)
    {
        var header = document.Header;
        foreach (var row in document.AsEnumerable())
        {
            if (asHashtable)
            {
                var rowValues = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
                for (var i = 0; i < header.Count && i < row.FieldCount; i++)
                {
                    rowValues.Add(header[i], row[i]);
                }

                writeObject(rowValues);
                continue;
            }

            var outputHeader = GetOutputHeader(header);
            var psObj = new PSObject();
            var valueCount = outputHeader.Length < row.FieldCount ? outputHeader.Length : row.FieldCount;
            var prevalidated = _prevalidatedOutputProperties;
            for (var i = 0; i < valueCount; i++)
            {
                psObj.Properties.Add(new PSNoteProperty(outputHeader[i], row[i]), prevalidated);
            }

            writeObject(psObj);
        }
    }

    public void WriteRow(IReadOnlyList<string> header, IReadOnlyList<string> row, bool asHashtable, Action<object?> writeObject)
    {
        var headerCount = header.Count;
        var rowCount = row.Count;
        var valueCount = rowCount < headerCount ? rowCount : headerCount;

        if (asHashtable)
        {
            var rowValues = new Dictionary<string, object?>(valueCount, StringComparer.OrdinalIgnoreCase);
            for (var i = 0; i < valueCount; i++)
            {
                rowValues.Add(header[i], row[i]);
            }

            writeObject(rowValues);
            return;
        }

        var outputHeader = GetOutputHeader(header);
        var psObj = new PSObject();
        var prevalidated = _prevalidatedOutputProperties;
        for (var i = 0; i < valueCount; i++)
        {
            psObj.Properties.Add(new PSNoteProperty(outputHeader[i], row[i]), prevalidated);
        }

        writeObject(psObj);
    }

    private string[] GetOutputHeader(IReadOnlyList<string> header)
    {
        if (_outputHeader is not null)
        {
            return _outputHeader;
        }

        var outputHeader = new string[header.Count];
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var canPrevalidate = true;
        for (var i = 0; i < header.Count; i++)
        {
            var name = header[i] ?? string.Empty;
            outputHeader[i] = name;

            if (string.IsNullOrEmpty(name))
            {
                canPrevalidate = false;
                continue;
            }

            if (!seen.Add(name))
            {
                throw new ExtendedTypeSystemException($"The member '{name}' is already present.");
            }
        }

        _prevalidatedOutputProperties = canPrevalidate;
        _outputHeader = outputHeader;
        return outputHeader;
    }
}
