using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Management.Automation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Creates a table from PowerShell objects.</summary>
/// <para>Transforms objects into an OfficeIMO table, applies styles/layout, and runs nested DSL customizations.</para>
/// <example>
///   <summary>Create a styled grid.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Add-OfficeWordTable -InputObject $Data -Style 'GridTable1LightAccent1' { WordTableCondition -FilterScript { $_.Total -gt 1000 } }</code>
///   <para>Writes a grid table and highlights rows exceeding $1,000.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeWordTable")]
[Alias("WordTable")]
public sealed class AddOfficeWordTableCommand : PSCmdlet
{
    /// <summary>Input data (array, list, DataTable, etc.).</summary>
    [Parameter(Mandatory = true, Position = 0)]
    [Alias("Data")]
    public object? InputObject { get; set; }

    /// <summary>Built-in table style.</summary>
    [Parameter]
    public WordTableStyle Style { get; set; } = WordTableStyle.TableGrid;

    /// <summary>Table layout behavior.</summary>
    [Parameter]
    [ValidateSet("Autofit", "Fixed", "AutoFitToContents", "AutoFitToWindow")]
    public string? Layout { get; set; }

    /// <summary>Skip writing header row.</summary>
    [Parameter]
    public SwitchParameter SkipHeader { get; set; }

    /// <summary>Transpose rows into property-oriented output.</summary>
    [Parameter]
    public SwitchParameter Transpose { get; set; }

    /// <summary>DSL content executed inside the table.</summary>
    [Parameter(Position = 1)]
    public ScriptBlock? Content { get; set; }

    /// <summary>Emit the created <see cref="WordTable"/>.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var rows = NormalizeRows(InputObject);
        if (rows.Length == 0)
        {
            ThrowTerminatingError(new ErrorRecord(
                new ArgumentException("Data cannot be empty.", nameof(InputObject)),
                "WordTableEmptyData",
                ErrorCategory.InvalidArgument,
                null));
            return;
        }

        var context = WordDslContext.Require(this);
        var tableRows = Transpose.IsPresent ? TransposeRows(rows) : rows;
        var normalizedRows = PowerShellObjectNormalizer.NormalizeItems(tableRows);
        var legacyLayout = ResolveLegacyLayout(Layout);
        var table = CreateTable(context, normalizedRows, Style, includeHeader: !SkipHeader.IsPresent, layout: legacyLayout);
        ApplyLayout(table, Layout);
        context.RegisterTableSource(table, tableRows);

        using (context.Push(table))
        {
            Content?.InvokeReturnAsIs();
        }

        var conditions = context.ConsumeTableConditions(table);
        if (conditions.Count > 0)
        {
            ApplyConditions(table, tableRows, conditions, SkipHeader.IsPresent);
        }

        context.ClearTableSource(table);

        if (PassThru.IsPresent)
        {
            WriteObject(table);
        }
    }

    private static object[] NormalizeRows(object? data)
    {
        switch (data)
        {
            case null:
                return Array.Empty<object>();
            case Array array:
                return array.Cast<object?>().Where(o => o != null).Select(o => o!).ToArray();
            case IEnumerable enumerable when data is not string:
                var list = new List<object>();
                foreach (var item in enumerable)
                {
                    if (item != null)
                    {
                        list.Add(item);
                    }
                }
                return list.ToArray();
            default:
                return new[] { data };
        }
    }

    private static object[] TransposeRows(IReadOnlyList<object> rows)
    {
        var maps = rows.Select(BuildPropertyMap).ToList();
        var columnNames = new List<string>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var map in maps)
        {
            foreach (var key in map.Keys)
            {
                if (seen.Add(key))
                {
                    columnNames.Add(key);
                }
            }
        }

        if (columnNames.Count == 0)
        {
            return Array.Empty<object>();
        }

        var transposed = new object[columnNames.Count];
        for (var columnIndex = 0; columnIndex < columnNames.Count; columnIndex++)
        {
            var propertyName = columnNames[columnIndex];
            var transposedRow = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase)
            {
                ["Property"] = propertyName
            };

            for (var rowIndex = 0; rowIndex < maps.Count; rowIndex++)
            {
                var header = $"Row{rowIndex + 1}";
                transposedRow[header] = maps[rowIndex].TryGetValue(propertyName, out var value)
                    ? value
                    : null;
            }

            transposed[columnIndex] = transposedRow;
        }

        return transposed;
    }

    private static Dictionary<string, object?> BuildPropertyMap(object row)
    {
        var normalized = PowerShellObjectNormalizer.NormalizeItem(row);
        if (normalized is IDictionary dictionary)
        {
            var mapped = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
            foreach (DictionaryEntry entry in dictionary)
            {
                var key = entry.Key?.ToString();
                if (string.IsNullOrWhiteSpace(key))
                {
                    continue;
                }

                var propertyName = key!;
                if (mapped.ContainsKey(propertyName))
                {
                    continue;
                }

                mapped[propertyName] = entry.Value;
            }

            if (mapped.Count > 0)
            {
                return mapped;
            }
        }

        return new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase)
        {
            ["Value"] = normalized
        };
    }

    private void ApplyConditions(WordTable table, IReadOnlyList<object> rows, IReadOnlyList<WordTableConditionModel> conditions, bool skipHeader)
    {
        var dataRowOffset = skipHeader ? 0 : 1;
        for (var index = 0; index < rows.Count && (index + dataRowOffset) < table.RowsCount; index++)
        {
            var rowObject = rows[index];
            var wordRow = table.Rows[index + dataRowOffset];
            foreach (var condition in conditions)
            {
                if (!EvaluateCondition(condition.FilterScript, rowObject))
                {
                    continue;
                }

                if (condition.TableStyle.HasValue)
                {
                    table.Style = condition.TableStyle.Value;
                }

                if (condition.BackgroundColor is string backgroundColor &&
                    !string.IsNullOrWhiteSpace(backgroundColor))
                {
                    foreach (var cell in wordRow.Cells)
                    {
                        cell.ShadingFillColorHex = backgroundColor;
                    }
                }
            }
        }
    }

    private bool EvaluateCondition(ScriptBlock filter, object row)
    {
        var variables = new List<PSVariable>
        {
            new PSVariable("_", row),
            new PSVariable("PSItem", row)
        };

        Collection<PSObject> result = filter.InvokeWithContext(null, variables, Array.Empty<object>());
        if (result.Count == 0)
        {
            return false;
        }

        return LanguagePrimitives.IsTrue(result[result.Count - 1]);
    }

    private static TableLayoutValues? ResolveLegacyLayout(string? layout)
    {
        if (string.IsNullOrWhiteSpace(layout))
        {
            return null;
        }

        if (string.Equals(layout, "AutoFitToContents", StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }

        if (string.Equals(layout, "AutoFitToWindow", StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }

        return string.Equals(layout, "Autofit", StringComparison.OrdinalIgnoreCase)
            ? TableLayoutValues.Autofit
            : TableLayoutValues.Fixed;
    }

    private static void ApplyLayout(WordTable table, string? layout)
    {
        if (string.IsNullOrWhiteSpace(layout))
        {
            return;
        }

        if (string.Equals(layout, "AutoFitToContents", StringComparison.OrdinalIgnoreCase))
        {
            table.LayoutMode = WordTableLayoutType.AutoFitToContents;
            return;
        }

        if (string.Equals(layout, "AutoFitToWindow", StringComparison.OrdinalIgnoreCase))
        {
            table.LayoutMode = WordTableLayoutType.AutoFitToWindow;
        }
    }

    private static WordTable CreateTable(
        WordDslContext context,
        IReadOnlyList<object?> normalizedRows,
        WordTableStyle style,
        bool includeHeader,
        TableLayoutValues? layout)
    {
        if (context.CurrentTableCell == null)
        {
            return context.Document.AddTableFromObjects(normalizedRows, style, includeHeader, layout);
        }

        return AddTableToCell(context.CurrentTableCell, normalizedRows, style, includeHeader, layout);
    }

    private static WordTable AddTableToCell(
        WordTableCell cell,
        IReadOnlyList<object?> items,
        WordTableStyle style,
        bool includeHeader,
        TableLayoutValues? layout)
    {
        if (items.Count == 0)
        {
            throw new ArgumentException("Provide at least one data row.", nameof(items));
        }

        var first = items[0];
        if (first == null)
        {
            throw new ArgumentException("Data rows cannot be null.", nameof(items));
        }

        var columns = GetColumnNames(first);
        if (columns.Count == 0)
        {
            throw new InvalidOperationException("Unable to infer column names. Use objects with properties or dictionaries.");
        }

        var rowCount = items.Count + (includeHeader ? 1 : 0);
        var table = cell.AddTable(rowCount, columns.Count, style);
        if (layout.HasValue)
        {
            table.LayoutType = layout.Value;
        }

        var rowIndex = 0;
        if (includeHeader)
        {
            for (var columnIndex = 0; columnIndex < columns.Count; columnIndex++)
            {
                SetCellText(table, 0, columnIndex, columns[columnIndex]);
            }

            rowIndex = 1;
        }

        for (var sourceIndex = 0; sourceIndex < items.Count; sourceIndex++)
        {
            var row = items[sourceIndex];
            if (row == null)
            {
                throw new InvalidOperationException("Data rows cannot contain null entries.");
            }

            for (var columnIndex = 0; columnIndex < columns.Count; columnIndex++)
            {
                var value = GetValue(row, columns[columnIndex]);
                SetCellText(table, rowIndex, columnIndex, value?.ToString() ?? string.Empty);
            }

            rowIndex++;
        }

        return table;
    }

    private static IReadOnlyList<string> GetColumnNames(object item)
    {
        if (item is IDictionary dictionary)
        {
            return dictionary.Keys
                .Cast<object?>()
                .Select(key => key?.ToString())
                .Where(name => !string.IsNullOrWhiteSpace(name))
                .Select(name => name!)
                .ToList();
        }

        return PSObject.AsPSObject(item).Properties
            .Where(property => property.MemberType == PSMemberTypes.NoteProperty || property.MemberType == PSMemberTypes.Property)
            .Select(property => property.Name)
            .Where(name => !string.IsNullOrWhiteSpace(name))
            .ToList();
    }

    private static object? GetValue(object item, string columnName)
    {
        if (item is IDictionary dictionary)
        {
            foreach (DictionaryEntry entry in dictionary)
            {
                if (string.Equals(entry.Key?.ToString(), columnName, StringComparison.OrdinalIgnoreCase))
                {
                    return entry.Value;
                }
            }

            return null;
        }

        return PSObject.AsPSObject(item).Properties[columnName]?.Value;
    }

    private static void SetCellText(WordTable table, int rowIndex, int columnIndex, string value)
    {
        var cell = table.Rows[rowIndex].Cells[columnIndex];
        var paragraph = cell.Paragraphs.Count > 0 ? cell.Paragraphs[0] : cell.AddParagraph();
        paragraph.Text = value;
    }
}
