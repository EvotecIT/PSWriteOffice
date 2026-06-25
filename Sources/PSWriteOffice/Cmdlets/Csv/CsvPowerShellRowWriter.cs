using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using System.Management.Automation;
using System.Reflection;
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
            var valueCount = outputHeader.Length < row.FieldCount ? outputHeader.Length : row.FieldCount;
            var psObj = PowerShellObjectFactory.Create(valueCount);
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
        var psObj = PowerShellObjectFactory.Create(valueCount);
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

    private static class PowerShellObjectFactory
    {
        // PowerShell 7+ exposes PSObject(int) for initial member capacity, but PowerShellStandard also exposes
        // PSObject(object). Use reflection so older runtimes fall back without wrapping the integer as BaseObject.
        private static readonly Func<int, PSObject>? CapacityFactory = CreateCapacityFactory();

        public static PSObject Create(int capacity)
        {
            return CapacityFactory?.Invoke(capacity) ?? new PSObject();
        }

        private static Func<int, PSObject>? CreateCapacityFactory()
        {
            var constructor = typeof(PSObject).GetConstructor(
                BindingFlags.Public | BindingFlags.Instance,
                binder: null,
                types: new[] { typeof(int) },
                modifiers: null);
            if (constructor is null)
            {
                return null;
            }

            try
            {
                var capacity = Expression.Parameter(typeof(int), "capacity");
                return Expression.Lambda<Func<int, PSObject>>(
                    Expression.New(constructor, capacity),
                    capacity).Compile();
            }
            catch
            {
                return capacity => (PSObject)constructor.Invoke(new object[] { capacity });
            }
        }
    }
}
