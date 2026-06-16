using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Management.Automation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Table;
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
    private readonly List<object?> _items = new();

    /// <summary>Input data (array, list, DataTable, etc.).</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true)]
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
    [Alias("SkipHeader")]
    public SwitchParameter NoHeader { get; set; }

    /// <summary>Projection to apply before writing the table.</summary>
    [Parameter]
    public OfficeTableView View { get; set; } = OfficeTableView.Normal;

    /// <summary>Legacy switch that maps to <see cref="OfficeTableView.Transpose"/>.</summary>
    [Parameter]
    public SwitchParameter Transpose { get; set; }

    /// <summary>DSL content executed inside the table.</summary>
    [Parameter(Position = 1)]
    public ScriptBlock? Content { get; set; }

    /// <summary>Existing paragraph after which the table should be inserted.</summary>
    [Parameter]
    public WordParagraph? AfterParagraph { get; set; }

    /// <summary>Existing paragraph before which the table should be inserted.</summary>
    [Parameter]
    public WordParagraph? BeforeParagraph { get; set; }

    /// <summary>Emit the created <see cref="WordTable"/>.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        TableInputCollector.AddInput(_items, InputObject);
    }

    /// <inheritdoc />
    protected override void EndProcessing()
    {
        if (AfterParagraph != null && BeforeParagraph != null)
        {
            throw new PSArgumentException("Use either -AfterParagraph or -BeforeParagraph, not both.");
        }

        var rows = TableInputCollector.RequireRows(_items, nameof(InputObject));
        if (rows.Length == 0)
        {
            ThrowTerminatingError(new ErrorRecord(
                new ArgumentException("Data cannot be empty.", nameof(InputObject)),
                "WordTableEmptyData",
                ErrorCategory.InvalidArgument,
                null));
            return;
        }

        var context = WordDslContext.Current;
        if (Content != null && context == null)
        {
            throw new InvalidOperationException("Nested Word table content must run inside a Word DSL script block.");
        }

        var effectiveView = Transpose.IsPresent ? OfficeTableView.Transpose : View;
        var tableRows = TableViewProjection.Project(rows, effectiveView);
        var normalizedRows = PowerShellObjectNormalizer.NormalizeItems(tableRows);
        var legacyLayout = ResolveLegacyLayout(Layout);
        var table = CreateTable(context, normalizedRows, Style, includeHeader: !NoHeader.IsPresent, layout: legacyLayout);
        ApplyLayout(table, Layout);

        if (context == null)
        {
            if (PassThru.IsPresent)
            {
                WriteObject(table);
            }

            return;
        }

        context.RegisterTableSource(table, tableRows);

        using (context.Push(table))
        {
            Content?.InvokeReturnAsIs();
        }

        var conditions = context.ConsumeTableConditions(table);
        if (conditions.Count > 0)
        {
            ApplyConditions(table, tableRows, conditions, NoHeader.IsPresent);
        }

        context.ClearTableSource(table);

        if (PassThru.IsPresent)
        {
            WriteObject(table);
        }
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

    private WordTable CreateTable(
        WordDslContext? context,
        IReadOnlyList<object?> normalizedRows,
        WordTableStyle style,
        bool includeHeader,
        TableLayoutValues? layout)
    {
        if (AfterParagraph != null)
        {
            return AddTableToHost(AfterParagraph, normalizedRows, style, includeHeader, layout);
        }

        if (BeforeParagraph != null)
        {
            return AddTableToHost(new TableInsertionAnchor(BeforeParagraph, true), normalizedRows, style, includeHeader, layout);
        }

        var host = context?.RequireParagraphHost()
            ?? throw new InvalidOperationException("Add-OfficeWordTable must run inside a Word DSL script block or use -AfterParagraph/-BeforeParagraph.");

        return AddTableToDslHost(context, host, normalizedRows, style, includeHeader, layout, context.CurrentParagraph);
    }

    private static WordTable AddTableToDslHost(
        WordDslContext context,
        object host,
        IReadOnlyList<object?> items,
        WordTableStyle style,
        bool includeHeader,
        TableLayoutValues? layout,
        WordParagraph? anchorParagraph)
    {
        var cursor = anchorParagraph != null
            ? anchorParagraph.AddParagraphAfterSelf()
            : context.AddParagraphToHost(host);
        var table = AddTableToHost(new TableInsertionAnchor(cursor, true), items, style, includeHeader, layout);
        context.RegisterBlockCursor(host, cursor);
        return table;
    }

    private static WordTable AddTableToHost(
        object host,
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
        var table = CreateEmptyTable(host, rowCount, columns.Count, style);
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

    private static WordTable CreateEmptyTable(object host, int rows, int columns, WordTableStyle style)
    {
        return host switch
        {
            WordTableCell cell => cell.AddTable(rows, columns, style),
            WordSection section => section.AddTable(rows, columns, style),
            WordHeader header => header.AddTable(rows, columns, style),
            WordFooter footer => footer.AddTable(rows, columns, style),
            WordParagraph paragraph => paragraph.AddTableAfter(rows, columns, style),
            TableInsertionAnchor anchor => anchor.Before
                ? anchor.Paragraph.AddTableBefore(rows, columns, style)
                : anchor.Paragraph.AddTableAfter(rows, columns, style),
            _ => throw new InvalidOperationException("Tables can only be added inside sections, headers, footers, table cells, or next to an existing paragraph.")
        };
    }

    private sealed record TableInsertionAnchor(WordParagraph Paragraph, bool Before);

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
