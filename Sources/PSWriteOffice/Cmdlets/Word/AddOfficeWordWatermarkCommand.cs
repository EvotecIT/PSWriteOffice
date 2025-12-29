using System;
using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Adds a watermark to the current section or header.</summary>
/// <para>Supports text or image watermarks using OfficeIMO.Word.</para>
/// <example>
///   <summary>Add a text watermark.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Add-OfficeWordWatermark -Text 'CONFIDENTIAL'</code>
///   <para>Inserts a text watermark into the current section.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeWordWatermark", DefaultParameterSetName = ParameterSetText)]
[Alias("WordWatermark")]
public sealed class AddOfficeWordWatermarkCommand : PSCmdlet
{
    private const string ParameterSetText = "Text";
    private const string ParameterSetImage = "Image";

    /// <summary>Watermark text.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetText)]
    public string Text { get; set; } = string.Empty;

    /// <summary>Path to an image watermark.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetImage)]
    public string ImagePath { get; set; } = string.Empty;

    /// <summary>Horizontal offset for the watermark.</summary>
    [Parameter]
    public double? HorizontalOffset { get; set; }

    /// <summary>Vertical offset for the watermark.</summary>
    [Parameter]
    public double? VerticalOffset { get; set; }

    /// <summary>Scale factor for the watermark.</summary>
    [Parameter]
    public double Scale { get; set; } = 1.0;

    /// <summary>Emit the created watermark.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var context = WordDslContext.Require(this);
        var section = context.RequireSection();
        var header = context.CurrentHeader;

        WordWatermark watermark;
        if (ParameterSetName == ParameterSetImage)
        {
            watermark = header != null
                ? header.AddWatermark(WordWatermarkStyle.Image, ImagePath, HorizontalOffset, VerticalOffset, Scale)
                : section.AddWatermark(WordWatermarkStyle.Image, ImagePath, HorizontalOffset, VerticalOffset, Scale);
        }
        else
        {
            watermark = header != null
                ? header.AddWatermark(WordWatermarkStyle.Text, Text, HorizontalOffset, VerticalOffset, Scale)
                : section.AddWatermark(WordWatermarkStyle.Text, Text, HorizontalOffset, VerticalOffset, Scale);
        }

        if (PassThru.IsPresent)
        {
            WriteObject(watermark);
        }
    }
}
