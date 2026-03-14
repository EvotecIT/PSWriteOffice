using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Markdown;
using PSWriteOffice.Services.Markdown;

namespace PSWriteOffice.Cmdlets.Markdown;

/// <summary>Adds a Markdown table from objects.</summary>
/// <example>
///   <summary>Add a table from input objects.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>MarkdownTable -InputObject $rows</code>
///   <para>Appends a Markdown table using the supplied objects.</para>
/// </example>
/// <example>
///   <summary>Append multiple tables to the same document.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$doc = New-OfficeMarkdown -Path .\Report.md -NoSave -PassThru
///   $doc | MarkdownTable -InputObject $summary -PassThru | MarkdownTable -InputObject $details</code>
///   <para>Creates two tables in sequence within the same Markdown document.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeMarkdownTable", DefaultParameterSetName = ParameterSetContext)]
[Alias("MarkdownTable")]
[OutputType(typeof(MarkdownDoc))]
public sealed class AddOfficeMarkdownTableCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";
    private readonly List<object?> _items = new();
    private MarkdownDoc? _document;

    /// <summary>Markdown document to update outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public MarkdownDoc Document { get; set; } = null!;

    /// <summary>Objects to convert into a Markdown table.</summary>
    [Parameter(ValueFromPipeline = true)]
    public object? InputObject { get; set; }

    /// <summary>Disable automatic alignment heuristics for tables.</summary>
    [Parameter]
    public SwitchParameter DisableAutoAlign { get; set; }

    /// <summary>Emit the Markdown document after appending the table.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void BeginProcessing()
    {
        _document = ResolveDocument();
    }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        AddInput(InputObject);
    }

    /// <inheritdoc />
    protected override void EndProcessing()
    {
        if (_items.Count == 0)
        {
            return;
        }

        var doc = _document ?? ResolveDocument();
        if (DisableAutoAlign.IsPresent)
        {
            doc.TableFrom(_items);
        }
        else
        {
            doc.TableFromAuto(_items);
        }

        if (PassThru.IsPresent)
        {
            WriteObject(doc);
        }
    }

    private void AddInput(object? value)
    {
        if (value is IEnumerable enumerable and not string and not IDictionary)
        {
            foreach (var entry in enumerable)
            {
                _items.Add(NormalizeItem(entry));
            }
            return;
        }

        _items.Add(NormalizeItem(value));
    }

    private static object? NormalizeItem(object? item)
    {
        if (item == null)
        {
            return null;
        }

        if (IsScalar(item))
        {
            return item;
        }

        var ps = PSObject.AsPSObject(item);
        if (ps.BaseObject is IDictionary dict)
        {
            return dict;
        }

        var properties = ps.Properties
            .Where(p => p.MemberType == PSMemberTypes.NoteProperty || p.MemberType == PSMemberTypes.Property)
            .Select(p => p.Name)
            .Where(n => !string.IsNullOrWhiteSpace(n))
            .ToList();

        if (properties.Count > 0)
        {
            var result = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
            foreach (var name in properties)
            {
                result[name] = ps.Properties[name]?.Value;
            }
            return result;
        }

        return item;
    }

    private static bool IsScalar(object item)
    {
        var type = item.GetType();
        return type.IsPrimitive
            || item is string
            || item is decimal
            || item is DateTime
            || item is DateTimeOffset
            || item is Guid;
    }

    private MarkdownDoc ResolveDocument()
    {
        if (ParameterSetName == ParameterSetDocument)
        {
            return Document ?? throw new PSArgumentException("Provide a Markdown document.");
        }

        return MarkdownDslContext.Require(this).Document;
    }
}
