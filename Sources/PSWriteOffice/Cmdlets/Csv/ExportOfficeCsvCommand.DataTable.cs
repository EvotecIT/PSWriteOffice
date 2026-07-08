using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Management.Automation;
using OfficeIMO.CSV;

namespace PSWriteOffice.Cmdlets.Csv;

public sealed partial class ExportOfficeCsvCommand
{
    private void ExportDataTable(DataTable table)
    {
        if (table == null || !TryPrepareOutput("Write CSV", allowAdditionalAppend: Append.IsPresent))
        {
            return;
        }

        if (Append.IsPresent)
        {
            AppendDataTable(table);
        }
        else
        {
            var options = CreateSaveOptions();
            using var writer = CreateTextWriter(append: false, options);
            WriteDataTable(writer, table, options);
        }

        _wroteOutput = true;
        WritePassThru();
    }

    private static bool TryGetDataTable(object? value, out DataTable table)
    {
        if (value is DataTable dataTable)
        {
            table = dataTable;
            return true;
        }

        if (value is PSObject { BaseObject: DataTable psObjectTable })
        {
            table = psObjectTable;
            return true;
        }

        table = null!;
        return false;
    }

    private static bool TryGetDataView(object? value, out DataView view)
    {
        if (value is DataView dataView)
        {
            view = dataView;
            return true;
        }

        if (value is PSObject { BaseObject: DataView psObjectView })
        {
            view = psObjectView;
            return true;
        }

        view = null!;
        return false;
    }

    private void AppendDataTable(DataTable table)
    {
        var options = CreateSaveOptions(includeHeader: !NoHeader.IsPresent && !_appendToExistingFile);
        var appendHeader = GetEffectiveAppendHeader(table);

        if (appendHeader is { Length: > 0 })
        {
            ValidateDataTableAppendHeader(table, appendHeader);
        }

        using var writer = CreateTextWriter(append: true, options);
        if (appendHeader is { Length: > 0 })
        {
            using var csvWriter = new CsvObjectWriter(writer, options);
            WriteDataTableRows(table, csvWriter, appendHeader);
            return;
        }

        WriteDataTable(writer, table, options);
    }

    private static void WriteDataTable(TextWriter writer, DataTable table, CsvSaveOptions options)
    {
        using var csvWriter = new CsvObjectWriter(writer, options, leaveOpen: true);
        using var reader = table.CreateDataReader();
        csvWriter.WriteDataReader(reader);
    }

    private string[]? GetEffectiveAppendHeader(DataTable table)
    {
        if (_appendHeader is not { Length: > 0 })
        {
            return null;
        }

        return _appendHeader;
    }

    private void ValidateDataTableAppendHeader(DataTable table, IReadOnlyList<string> appendHeader)
    {
        if (Force.IsPresent)
        {
            return;
        }

        foreach (var column in appendHeader)
        {
            if (!ContainsDataColumn(table, column))
            {
                throw new CsvException($"Cannot append CSV because the DataTable is missing the existing column '{column}'. Use -Force to append with blank values for missing columns.");
            }
        }
    }

    private static void WriteDataTableRows(DataTable table, CsvObjectWriter writer, IReadOnlyList<string> columns)
    {
        foreach (DataRow row in table.Rows)
        {
            writer.WriteRow(
                columns,
                columns.Count,
                (Row: row, Columns: columns),
                static (state, index) => TryGetDataTableValue(state.Row, state.Columns[index]));
        }
    }

    private static object? TryGetDataTableValue(DataRow row, string column)
    {
        var dataColumn = GetDataColumn(row.Table, column);
        if (dataColumn == null)
        {
            return null;
        }

        var value = row[dataColumn];
        return value == DBNull.Value ? null : value;
    }

    private static bool ContainsDataColumn(DataTable table, string column) => GetDataColumn(table, column) != null;

    private static DataColumn? GetDataColumn(DataTable table, string column)
    {
        if (table.Columns.Contains(column))
        {
            return table.Columns[column];
        }

        foreach (DataColumn dataColumn in table.Columns)
        {
            if (string.Equals(dataColumn.ColumnName, column, StringComparison.OrdinalIgnoreCase))
            {
                return dataColumn;
            }
        }

        return null;
    }
}
