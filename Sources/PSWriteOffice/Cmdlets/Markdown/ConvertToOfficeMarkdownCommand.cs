using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Markdown;

namespace PSWriteOffice.Cmdlets.Markdown;

/// <summary>Converts objects into a Markdown table.</summary>
/// <para>Returns Markdown text by default; use <c>-PassThru</c> to emit a <see cref="MarkdownDoc"/>.</para>
/// <example>
///   <summary>Convert objects to Markdown table.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$markdown = $data | ConvertTo-OfficeMarkdown</code>
///   <para>Generates Markdown table text from the input objects.</para>
/// </example>
[Cmdlet(VerbsData.ConvertTo, "OfficeMarkdown")]
[OutputType(typeof(string), typeof(MarkdownDoc))]
public sealed class ConvertToOfficeMarkdownCommand : PSCmdlet
{
    private readonly List<object?> _items = new();

    /// <summary>Objects to convert into Markdown.</summary>
    [Parameter(ValueFromPipeline = true)]
    public object? InputObject { get; set; }

    /// <summary>Disable automatic alignment heuristics for tables.</summary>
    [Parameter]
    public SwitchParameter DisableAutoAlign { get; set; }

    /// <summary>Emit a Markdown document object instead of text.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        _items.Add(NormalizeItem(InputObject));
    }

    /// <inheritdoc />
    protected override void EndProcessing()
    {
        if (_items.Count == 0)
        {
            return;
        }

        var doc = MarkdownDoc.Create();
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
        else
        {
            WriteObject(doc.ToMarkdown());
        }
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
}
