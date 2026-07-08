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
using PSWriteOffice.Services.Text;
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

        var context = WordDslContext.Require(this);
        var effectiveView = Transpose.IsPresent ? OfficeTableView.Transpose : View;
        var tableRows = TableViewProjection.Project(rows, effectiveView);
        var legacyLayout = ResolveLegacyLayout(Layout);
        int? conditionHeaderRowIndex = NoHeader.IsPresent ? null : 0;
        WordTable table;
        if (OfficeTableSpecParser.TryCreate(
                tableRows,
                propertyNames: null,
                header: NoHeader.IsPresent ? Array.Empty<string>() : null,
                out var tableSpec))
        {
            table = CreateTable(context, tableSpec, Style, legacyLayout);
            conditionHeaderRowIndex = tableSpec.HeaderRowIndex;
        }
        else
        {
            table = CreateTable(
                context,
                PowerShellObjectNormalizer.NormalizeItems(tableRows),
                Style,
                includeHeader: !NoHeader.IsPresent,
                layout: legacyLayout);
        }

        ApplyLayout(table, Layout);
        context.RegisterTableSource(table, tableRows);

        using (context.Push(table))
        {
            Content?.InvokeReturnAsIs();
        }

        var conditions = context.ConsumeTableConditions(table);
        if (conditions.Count > 0)
        {
            ApplyConditions(table, tableRows, conditions, conditionHeaderRowIndex);
        }

        context.ClearTableSource(table);

        if (PassThru.IsPresent)
        {
            WriteObject(table);
        }
    }

    private void ApplyConditions(WordTable table, IReadOnlyList<object> rows, IReadOnlyList<WordTableConditionModel> conditions, int? headerRowIndex)
    {
        for (var index = 0; index < rows.Count; index++)
        {
            var targetRowIndex = headerRowIndex.HasValue && index >= headerRowIndex.Value
                ? index + 1
                : index;
            if (targetRowIndex >= table.RowsCount)
            {
                break;
            }

            var rowObject = rows[index];
            var wordRow = table.Rows[targetRowIndex];
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

    private static WordTable CreateTable(
        WordDslContext context,
        OfficeTableSpec spec,
        WordTableStyle style,
        TableLayoutValues? layout)
    {
        if (spec.RowCount == 0 || spec.ColumnCount == 0)
        {
            throw new ArgumentException("Provide at least one table row and one table column.", nameof(spec));
        }

        var table = context.CurrentTableCell == null
            ? context.Document.AddTable(spec.RowCount, spec.ColumnCount, style)
            : context.CurrentTableCell.AddTable(spec.RowCount, spec.ColumnCount, style);

        if (layout.HasValue)
        {
            table.LayoutType = layout.Value;
        }

        foreach (var placement in spec.Placements)
        {
            SetCell(table, placement.RowIndex, placement.ColumnIndex, placement.Cell);
        }

        ApplyCellSpans(table, spec.Placements);

        return table;
    }

    private static void ApplyCellSpans(WordTable table, IReadOnlyList<OfficeTableCellPlacement> placements)
    {
        var spanPlacements = placements
            .Where(static placement => placement.Cell.HasSpan)
            .ToList();

        foreach (var placement in spanPlacements
            .Where(static placement => placement.Cell.RowSpan > 1)
            .OrderBy(static placement => placement.RowIndex)
            .ThenByDescending(static placement => placement.ColumnIndex))
        {
            for (var offset = placement.Cell.ColumnSpan - 1; offset >= 0; offset--)
            {
                table.Rows[placement.RowIndex]
                    .Cells[placement.ColumnIndex + offset]
                    .MergeVertically(placement.Cell.RowSpan - 1);
            }
        }

        var horizontalSpans = spanPlacements
            .Where(static placement => placement.Cell.ColumnSpan > 1)
            .SelectMany(static placement => Enumerable
                .Range(placement.RowIndex, placement.Cell.RowSpan)
                .Select(rowIndex => new
                {
                    RowIndex = rowIndex,
                    placement.ColumnIndex,
                    placement.Cell.ColumnSpan
                }))
            .OrderByDescending(static span => span.RowIndex)
            .ThenByDescending(static span => span.ColumnIndex);

        foreach (var span in horizontalSpans)
        {
            table.Rows[span.RowIndex]
                .Cells[span.ColumnIndex]
                .MergeHorizontally(span.ColumnSpan - 1);
        }
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

    private static void SetCell(WordTable table, int rowIndex, int columnIndex, OfficeTableCellSpec spec)
    {
        var cell = table.Rows[rowIndex].Cells[columnIndex];
        var paragraph = cell.Paragraphs.Count > 0 ? cell.Paragraphs[0] : cell.AddParagraph();
        paragraph.Text = string.Empty;

        if (spec.HasRuns)
        {
            WordTextRunService.ApplyRuns(paragraph, spec.Runs!.ToArray());
        }
        else if (spec.Style?.HasTextStyle == true)
        {
            WordTextRunService.ApplyRuns(paragraph, new[]
            {
                new OfficeTextRunSpec
                {
                    Text = spec.Text,
                    Bold = spec.Style.Bold,
                    Italic = spec.Style.Italic,
                    Underline = spec.Style.Underline,
                    UnderlineStyle = spec.Style.UnderlineStyle,
                    Strike = spec.Style.Strike,
                    Color = spec.Style.TextColor,
                    FontSize = spec.Style.FontSize
                }
            });
        }
        else
        {
            paragraph.Text = spec.Text;
        }

        ApplyCellStyle(cell, paragraph, spec.Style);
    }

    private static void ApplyCellStyle(WordTableCell cell, WordParagraph paragraph, OfficeTableCellStyle? style)
    {
        if (style == null)
        {
            return;
        }

        var fill = OfficeColorUtilities.ToRgbHex(style.FillColor);
        if (!string.IsNullOrWhiteSpace(fill))
        {
            cell.ShadingFillColorHex = fill!;
        }

        if (!string.IsNullOrWhiteSpace(style.Align) &&
            OpenXmlValueParser.TryParse<JustificationValues>(style.Align, out var alignment))
        {
            paragraph.ParagraphAlignment = alignment;
        }

        if (!string.IsNullOrWhiteSpace(style.VerticalAlign) &&
            OpenXmlValueParser.TryParse<TableVerticalAlignmentValues>(style.VerticalAlign, out var verticalAlignment))
        {
            cell.VerticalAlignment = verticalAlignment;
        }
    }
}
