using System;
using System.Management.Automation;
using System.Reflection;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Gets layout placeholder metadata for a slide.</summary>
/// <example>
///   <summary>List layout placeholders.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Get-OfficePowerPointLayoutPlaceholder -Slide $slide</code>
///   <para>Returns the layout placeholder definitions for the slide.</para>
/// </example>
/// <example>
///   <summary>Inspect layout placeholders inside the DSL.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficePowerPoint -Path .\deck.pptx { PptSlide { Get-OfficePowerPointLayoutPlaceholder } }</code>
///   <para>Uses the current slide context.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficePowerPointLayoutPlaceholder")]
[Alias("PptLayoutPlaceholders")]
[OutputType(typeof(PowerPointLayoutPlaceholderInfo))]
public sealed class GetOfficePowerPointLayoutPlaceholderCommand : PSCmdlet
{
    /// <summary>Slide to inspect (optional inside DSL).</summary>
    [Parameter(ValueFromPipeline = true)]
    public PowerPointSlide? Slide { get; set; }

    /// <summary>Placeholder type to filter on.</summary>
    [Parameter]
    [Alias("Type")]
    public string? PlaceholderType { get; set; }

    /// <summary>Optional placeholder index.</summary>
    [Parameter]
    public uint? Index { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var slide = Slide ?? PowerPointDslContext.Require(this).RequireSlide();

        if (string.IsNullOrWhiteSpace(PlaceholderType))
        {
            WriteObject(slide.GetLayoutPlaceholders(), enumerateCollection: true);
            return;
        }

        if (!TryResolvePlaceholderType(PlaceholderType, out var placeholderType))
        {
            throw new PSArgumentException($"Unknown placeholder type '{PlaceholderType}'.", nameof(PlaceholderType));
        }

        var placeholder = slide.GetLayoutPlaceholder(placeholderType, Index);
        if (placeholder != null)
        {
            WriteObject(placeholder);
        }
    }

    private static bool TryResolvePlaceholderType(string? placeholderType, out PlaceholderValues value)
    {
        value = default;
        if (string.IsNullOrWhiteSpace(placeholderType))
        {
            return false;
        }

        var property = typeof(PlaceholderValues).GetProperty(
            placeholderType,
            BindingFlags.Public | BindingFlags.Static | BindingFlags.IgnoreCase);

        if (property == null)
        {
            return false;
        }

        value = (PlaceholderValues)property.GetValue(null)!;
        return true;
    }
}
