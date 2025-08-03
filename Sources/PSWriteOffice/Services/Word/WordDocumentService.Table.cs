using System;
using System.Collections;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace PSWriteOffice.Services.Word;

public static partial class WordDocumentService
{
    public static WordTable AddTable(WordDocument document, Array dataTable, WordTableStyle style, string? tableLayout,
        bool skipHeader)
    {
        if (dataTable.Length == 0)
        {
            throw new ArgumentException("DataTable is empty", nameof(dataTable));
        }

        string[] properties;
        if (dataTable.GetValue(0) is IDictionary)
        {
            properties = new[] { "Name", "Value" };
        }
        else
        {
            var obj = dataTable.GetValue(0)!;
            properties = obj.GetType().GetProperties().Select(p => p.Name).ToArray();
        }

        var rows = dataTable.Length;
        var columns = properties.Length;
        WordTable table;
        var rowIndex = 0;

        if (!skipHeader)
        {
            table = document.AddTable(rows + 1, columns, style);
            for (var c = 0; c < columns; c++)
            {
                table.Rows[0].Cells[c].Paragraphs[0].Text = properties[c];
            }
            rowIndex = 1;
        }
        else
        {
            table = document.AddTable(rows, columns, style);
        }

        if (!string.IsNullOrEmpty(tableLayout))
        {
            table.LayoutType = tableLayout.Equals("Autofit", StringComparison.OrdinalIgnoreCase)
                ? TableLayoutValues.Autofit
                : TableLayoutValues.Fixed;
        }

        for (var r = 0; r < rows; r++)
        {
            var rowObj = dataTable.GetValue(r)!;
            for (var c = 0; c < columns; c++)
            {
                object? value;
                if (rowObj is IDictionary dict)
                {
                    value = dict[properties[c]];
                }
                else
                {
                    value = rowObj.GetType().GetProperty(properties[c])?.GetValue(rowObj);
                }

                table.Rows[rowIndex].Cells[c].Paragraphs[0].Text = value?.ToString() ?? string.Empty;
            }

            rowIndex++;
        }

        return table;
    }
}
