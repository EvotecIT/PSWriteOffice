using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Management.Automation;
using OfficeIMO.CSV;

namespace PSWriteOffice.Cmdlets.Csv;

public sealed partial class ExportOfficeCsvCommand
{
    private void ExportDataReader(IDataReader reader)
    {
        if (reader == null || !TryPrepareOutput("Write CSV", allowAdditionalAppend: Append.IsPresent))
        {
            return;
        }

        if (Append.IsPresent)
        {
            AppendDataReader(reader);
        }
        else
        {
            var options = CreateSaveOptions();
            using var writer = CreateTextWriter(append: false, options);
            WriteDataReader(writer, reader, options);
        }

        _wroteOutput = true;
        WritePassThru();
    }

    private static bool TryGetDataReader(object? value, out IDataReader reader)
    {
        if (value is IDataReader dataReader)
        {
            reader = dataReader;
            return true;
        }

        if (value is PSObject { BaseObject: IDataReader psObjectReader })
        {
            reader = psObjectReader;
            return true;
        }

        reader = null!;
        return false;
    }

    private void AppendDataReader(IDataReader reader)
    {
        var options = CreateSaveOptions(includeHeader: !NoHeader.IsPresent && !_appendToExistingFile);
        var appendHeader = GetEffectiveAppendHeader(reader);

        if (appendHeader is { Length: > 0 })
        {
            ValidateDataReaderAppendHeader(reader, appendHeader);
        }

        using var writer = CreateTextWriter(append: true, options);
        if (appendHeader is { Length: > 0 })
        {
            using var csvWriter = new CsvObjectWriter(writer, options);
            WriteDataReaderRows(reader, csvWriter, appendHeader);
            return;
        }

        WriteDataReader(writer, reader, options);
    }

    private static void WriteDataReader(TextWriter writer, IDataReader reader, CsvSaveOptions options)
    {
        using var csvWriter = new CsvObjectWriter(writer, options, leaveOpen: true);
        csvWriter.WriteDataReader(reader);
    }

    private string[]? GetEffectiveAppendHeader(IDataReader reader)
    {
        if (_appendHeader is not { Length: > 0 })
        {
            return null;
        }

        if (!NoHeader.IsPresent || Force.IsPresent || DataReaderContainsColumns(reader, _appendHeader))
        {
            return _appendHeader;
        }

        return null;
    }

    private void ValidateDataReaderAppendHeader(IDataReader reader, IReadOnlyList<string> appendHeader)
    {
        if (Force.IsPresent)
        {
            return;
        }

        foreach (var column in appendHeader)
        {
            if (!ContainsDataReaderColumn(reader, column))
            {
                throw new CsvException($"Cannot append CSV because the data reader is missing the existing column '{column}'. Use -Force to append with blank values for missing columns.");
            }
        }
    }

    private static void WriteDataReaderRows(IDataReader reader, CsvObjectWriter writer, IReadOnlyList<string> columns)
    {
        var columnOrdinals = GetDataReaderColumnOrdinals(reader);
        while (reader.Read())
        {
            writer.WriteRow(
                columns,
                columns.Count,
                (Reader: reader, Columns: columns, Ordinals: columnOrdinals),
                static (state, index) => TryGetDataReaderValue(state.Reader, state.Columns[index], state.Ordinals));
        }
    }

    private static object? TryGetDataReaderValue(IDataRecord reader, string column, IReadOnlyDictionary<string, int> ordinals)
    {
        if (!ordinals.TryGetValue(column, out var ordinal))
        {
            return null;
        }

        var value = reader.GetValue(ordinal);
        return value == DBNull.Value ? null : value;
    }

    private static bool DataReaderContainsColumns(IDataReader reader, IEnumerable<string> columns)
    {
        foreach (var column in columns)
        {
            if (!ContainsDataReaderColumn(reader, column))
            {
                return false;
            }
        }

        return true;
    }

    private static bool ContainsDataReaderColumn(IDataReader reader, string column)
    {
        for (var i = 0; i < reader.FieldCount; i++)
        {
            if (string.Equals(reader.GetName(i), column, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }

        return false;
    }

    private static Dictionary<string, int> GetDataReaderColumnOrdinals(IDataReader reader)
    {
        var ordinals = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        for (var i = 0; i < reader.FieldCount; i++)
        {
            var name = reader.GetName(i);
            if (!ordinals.ContainsKey(name))
            {
                ordinals.Add(name, i);
            }
        }

        return ordinals;
    }
}
