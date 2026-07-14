using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Management.Automation;
using OfficeIMO.Excel;

namespace PSWriteOffice.Services.Excel;

internal static class ExcelReadOutputService
{
    public static ExcelReadOptions CreateOptions(bool numericAsDecimal, bool useCachedFormulaResult = true, CultureInfo? culture = null)
    {
        return new ExcelReadOptions {
            NumericAsDecimal = numericAsDecimal,
            UseCachedFormulaResult = useCachedFormulaResult,
            Culture = culture ?? CultureInfo.InvariantCulture
        };
    }

    public static ExcelSheetReader ResolveSheetReader(ExcelDocumentReader reader, string? sheetName, int? sheetIndex)
    {
        if (reader == null)
        {
            throw new ArgumentNullException(nameof(reader));
        }

        if (!string.IsNullOrWhiteSpace(sheetName) && sheetIndex.HasValue)
        {
            throw new PSArgumentException("Specify either -Sheet or -SheetIndex, but not both.");
        }

        if (!string.IsNullOrWhiteSpace(sheetName))
        {
            return reader.GetSheet(sheetName!);
        }

        if (sheetIndex.HasValue)
        {
            if (sheetIndex.Value < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(sheetIndex), "SheetIndex must be 0 or greater.");
            }

            return reader.GetSheet(sheetIndex.Value + 1);
        }

        if (reader.SheetCount == 0)
        {
            throw new InvalidOperationException("Workbook contains no worksheets.");
        }

        return reader.GetSheet(1);
    }

    public static void WriteOutput(
        IAsyncCmdletPipeline cmdlet,
        DataTable table,
        bool asDataTable,
        bool asHashtable,
        bool byColumn = false,
        string? worksheetName = null)
    {
        if (cmdlet == null)
        {
            throw new ArgumentNullException(nameof(cmdlet));
        }

        if (table == null)
        {
            throw new ArgumentNullException(nameof(table));
        }

        if (asDataTable && asHashtable)
        {
            throw new PSArgumentException("Specify either -AsDataTable or -AsHashtable, but not both.");
        }

        if (asDataTable && byColumn)
        {
            throw new PSArgumentException("Specify either -AsDataTable or -ByColumn, but not both.");
        }

        if (asDataTable)
        {
            if (!string.IsNullOrWhiteSpace(worksheetName))
            {
                table.TableName = worksheetName!;
                table.ExtendedProperties["WorksheetName"] = worksheetName!;
            }

            cmdlet.WriteObject(table, enumerateCollection: false);
            return;
        }

        var columnCount = table.Columns.Count;
        var columnNames = new string[columnCount];
        for (var i = 0; i < columnCount; i++)
        {
            columnNames[i] = table.Columns[i].ColumnName;
        }
        var rowOutputColumnNames = CreateRowOutputColumnNames(columnNames, !string.IsNullOrWhiteSpace(worksheetName));

        if (byColumn)
        {
            WriteColumnOutput(cmdlet, table, columnNames, asHashtable, worksheetName);
            return;
        }

        foreach (DataRow row in table.Rows)
        {
            if (asHashtable)
            {
                var hashtable = new Hashtable(columnCount + (worksheetName == null ? 0 : 1), StringComparer.OrdinalIgnoreCase);
                if (!string.IsNullOrWhiteSpace(worksheetName))
                {
                    hashtable["WorksheetName"] = worksheetName!;
                }

                for (var columnIndex = 0; columnIndex < columnCount; columnIndex++)
                {
                    var value = row[columnIndex];
                    hashtable[rowOutputColumnNames[columnIndex]] = value is DBNull ? null : value;
                }

                cmdlet.WriteObject(hashtable);
            }
            else
            {
                var psObject = new PSObject();
                if (!string.IsNullOrWhiteSpace(worksheetName))
                {
                    psObject.Properties.Add(new PSNoteProperty("WorksheetName", worksheetName!));
                }

                for (var columnIndex = 0; columnIndex < columnCount; columnIndex++)
                {
                    var value = row[columnIndex];
                    psObject.Properties.Add(new PSNoteProperty(rowOutputColumnNames[columnIndex], value is DBNull ? null : value));
                }

                cmdlet.WriteObject(psObject);
            }
        }
    }

    private static string[] CreateRowOutputColumnNames(string[] columnNames, bool hasWorksheetMetadata)
    {
        if (!hasWorksheetMetadata)
        {
            return columnNames;
        }

        var outputNames = new string[columnNames.Length];
        var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "WorksheetName"
        };

        for (var i = 0; i < columnNames.Length; i++)
        {
            var candidate = columnNames[i];
            if (used.Add(candidate))
            {
                outputNames[i] = candidate;
                continue;
            }

            outputNames[i] = CreateUniqueColumnName(candidate, used);
        }

        return outputNames;
    }

    private static string CreateUniqueColumnName(string columnName, ISet<string> used)
    {
        var baseName = string.IsNullOrWhiteSpace(columnName) ? "Column" : columnName;
        var candidate = $"{baseName}Value";
        var index = 2;
        while (!used.Add(candidate))
        {
            candidate = $"{baseName}Value{index++}";
        }

        return candidate;
    }

    private static void WriteColumnOutput(IAsyncCmdletPipeline cmdlet, DataTable table, string[] columnNames, bool asHashtable, string? worksheetName)
    {
        for (var columnIndex = 0; columnIndex < columnNames.Length; columnIndex++)
        {
            var values = new object?[table.Rows.Count];
            for (var rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++)
            {
                var value = table.Rows[rowIndex][columnIndex];
                values[rowIndex] = value is DBNull ? null : value;
            }

            if (asHashtable)
            {
                var hashtable = new Hashtable(StringComparer.OrdinalIgnoreCase)
                {
                    ["ColumnName"] = columnNames[columnIndex],
                    ["ColumnIndex"] = columnIndex + 1,
                    ["Values"] = values
                };
                if (!string.IsNullOrWhiteSpace(worksheetName))
                {
                    hashtable["WorksheetName"] = worksheetName!;
                }

                cmdlet.WriteObject(hashtable);
            }
            else
            {
                var psObject = new PSObject();
                if (!string.IsNullOrWhiteSpace(worksheetName))
                {
                    psObject.Properties.Add(new PSNoteProperty("WorksheetName", worksheetName!));
                }

                psObject.Properties.Add(new PSNoteProperty("ColumnName", columnNames[columnIndex]));
                psObject.Properties.Add(new PSNoteProperty("ColumnIndex", columnIndex + 1));
                psObject.Properties.Add(new PSNoteProperty("Values", values));
                cmdlet.WriteObject(psObject);
            }
        }
    }
}
