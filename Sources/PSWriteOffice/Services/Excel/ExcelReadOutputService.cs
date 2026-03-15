using System;
using System.Collections;
using System.Data;
using System.Globalization;
using System.Management.Automation;
using OfficeIMO.Excel;

namespace PSWriteOffice.Services.Excel;

internal static class ExcelReadOutputService
{
    public static ExcelReadOptions CreateOptions(bool numericAsDecimal)
    {
        return new ExcelReadOptions {
            NumericAsDecimal = numericAsDecimal,
            Culture = CultureInfo.InvariantCulture
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

        foreach (DataRow row in table.Rows)
        {
            if (asHashtable)
            {
                var hashtable = new Hashtable(StringComparer.OrdinalIgnoreCase);
                foreach (DataColumn column in table.Columns)
                {
                    var value = row[column];
                    hashtable[column.ColumnName] = value is DBNull ? null : value;
                }

                cmdlet.WriteObject(hashtable);
            }
            else
            {
                var psObject = new PSObject();
                foreach (DataColumn column in table.Columns)
                {
                    var value = row[column];
                    psObject.Properties.Add(new PSNoteProperty(column.ColumnName, value is DBNull ? null : value));
                }

                cmdlet.WriteObject(psObject);
            }
        }
    }
}
