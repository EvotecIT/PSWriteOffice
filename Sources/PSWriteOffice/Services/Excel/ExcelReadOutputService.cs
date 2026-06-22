using System;
using System.Collections;
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

    public static void WriteOutput(PSCmdlet cmdlet, DataTable table, bool asDataTable, bool asHashtable)
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

        if (asDataTable)
        {
            cmdlet.WriteObject(table, enumerateCollection: false);
            return;
        }

        var columnCount = table.Columns.Count;
        var columnNames = new string[columnCount];
        for (var i = 0; i < columnCount; i++)
        {
            columnNames[i] = table.Columns[i].ColumnName;
        }

        foreach (DataRow row in table.Rows)
        {
            if (asHashtable)
            {
                var hashtable = new Hashtable(columnCount, StringComparer.OrdinalIgnoreCase);
                for (var columnIndex = 0; columnIndex < columnCount; columnIndex++)
                {
                    var value = row[columnIndex];
                    hashtable[columnNames[columnIndex]] = value is DBNull ? null : value;
                }

                cmdlet.WriteObject(hashtable);
            }
            else
            {
                var psObject = new PSObject();
                for (var columnIndex = 0; columnIndex < columnCount; columnIndex++)
                {
                    var value = row[columnIndex];
                    psObject.Properties.Add(new PSNoteProperty(columnNames[columnIndex], value is DBNull ? null : value));
                }

                cmdlet.WriteObject(psObject);
            }
        }
    }
}
