using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Management.Automation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
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
    [ValidateSet("Autofit", "Fixed")]
    public string? Layout { get; set; }

    /// <summary>Skip writing header row.</summary>
    [Parameter]
    public SwitchParameter SkipHeader { get; set; }

    /// <summary>DSL content executed inside the table.</summary>
    [Parameter(Position = 1)]
    public ScriptBlock? Content { get; set; }

    /// <summary>Emit the created <see cref="WordTable"/>.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
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
        var table = WordDocumentService.AddTable(context.Document, rows, Style, Layout, SkipHeader.IsPresent);
        context.RegisterTableSource(table, rows);

        using (context.Push(table))
        {
            Content?.InvokeReturnAsIs();
        }

        var conditions = context.ConsumeTableConditions(table);
        if (conditions.Count > 0)
        {
            ApplyConditions(table, rows, conditions, SkipHeader.IsPresent);
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
}
