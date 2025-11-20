using System;
using System.Collections;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace PSWriteOffice.Services.Word;

public static partial class WordDocumentService
{
    /// <summary>Creates a table from an array of objects.</summary>
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
                var headerParagraph = GetOrCreateParagraph(table, 0, c);
                headerParagraph.Text = properties[c];
            }
            rowIndex = 1;
        }
        else
        {
            table = document.AddTable(rows, columns, style);
        }

        if (tableLayout is { Length: > 0 } layout)
        {
            table.LayoutType = layout.Equals("Autofit", StringComparison.OrdinalIgnoreCase)
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

                var paragraph = GetOrCreateParagraph(table, rowIndex, c);
                paragraph.Text = value?.ToString() ?? string.Empty;
            }

            rowIndex++;
        }

        return table;
    }

    private static WordParagraph GetOrCreateParagraph(WordTable table, int rowIndex, int columnIndex)
    {
        var rows = table.Rows ?? throw new InvalidOperationException("Table rows collection is missing.");
        var row = rows[rowIndex];
        var cells = row.Cells ?? throw new InvalidOperationException("Table cells collection is missing.");
        var cell = cells[columnIndex];
        return cell.Paragraphs.Count > 0 ? cell.Paragraphs[0] : cell.AddParagraph();
    }
}
