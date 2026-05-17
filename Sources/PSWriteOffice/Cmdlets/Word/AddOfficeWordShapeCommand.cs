using System;
using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Adds a basic OfficeIMO Word shape to the current paragraph.</summary>
[Cmdlet(VerbsCommon.Add, "OfficeWordShape")]
[Alias("WordShape")]
[OutputType(typeof(WordShape))]
public sealed class AddOfficeWordShapeCommand : PSCmdlet
{
    /// <summary>Shape type to add.</summary>
    [Parameter]
    public ShapeType Type { get; set; } = ShapeType.Rectangle;

    /// <summary>Width in points.</summary>
    [Parameter]
    [ValidateRange(0.1, double.MaxValue)]
    public double Width { get; set; } = 120;

    /// <summary>Height in points.</summary>
    [Parameter]
    [ValidateRange(0.1, double.MaxValue)]
    public double Height { get; set; } = 48;

    /// <summary>Anchored left position in points. When omitted, the shape is inline.</summary>
    [Parameter]
    [ValidateRange(0, double.MaxValue)]
    public double? Left { get; set; }

    /// <summary>Anchored top position in points. When omitted, the shape is inline.</summary>
    [Parameter]
    [ValidateRange(0, double.MaxValue)]
    public double? Top { get; set; }

    /// <summary>Fill color as #RRGGBB.</summary>
    [Parameter]
    public string? FillColor { get; set; }

    /// <summary>Stroke color as #RRGGBB.</summary>
    [Parameter]
    public string? StrokeColor { get; set; }

    /// <summary>Stroke width in points.</summary>
    [Parameter]
    [ValidateRange(0, double.MaxValue)]
    public double? StrokeWidth { get; set; }

    /// <summary>Optional title metadata.</summary>
    [Parameter]
    public string? Title { get; set; }

    /// <summary>Optional alternate text metadata.</summary>
    [Parameter]
    public string? Description { get; set; }

    /// <summary>Emit the created shape.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var context = WordDslContext.Require(this);
        var paragraph = context.CurrentParagraph ?? context.RequireParagraphHost().AddParagraph();

        var shape = Left.HasValue || Top.HasValue
            ? WordShape.AddDrawingShapeAnchored(paragraph, Type, Width, Height, Left ?? 0, Top ?? 0)
            : WordShape.AddDrawingShape(paragraph, Type, Width, Height);

        ApplyShape(shape);

        if (PassThru.IsPresent)
        {
            WriteObject(shape);
        }
    }

    private void ApplyShape(WordShape shape)
    {
        if (!string.IsNullOrWhiteSpace(FillColor))
        {
            shape.FillColorHex = FillColor!;
        }

        if (!string.IsNullOrWhiteSpace(StrokeColor))
        {
            shape.StrokeColorHex = StrokeColor!;
        }

        if (StrokeWidth.HasValue)
        {
            shape.StrokeWeight = StrokeWidth;
        }

        if (Title != null)
        {
            shape.Title = Title;
        }

        if (Description != null)
        {
            shape.Description = Description;
        }
    }
}
