using System;
using System.Management.Automation;
using System.Reflection;
using DocumentFormat.OpenXml.Drawing;
using OfficeIMO.PowerPoint;
using SixLabors.ImageSharp;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Adds a basic shape to a slide.</summary>
/// <para>Creates an auto shape at the requested coordinates and applies optional fill and outline styling.</para>
/// <example>
///   <summary>Create a rectangle highlight.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Add-OfficePowerPointShape -Slide $slide -ShapeType Rectangle -X 60 -Y 80 -Width 220 -Height 120 -FillColor '#DDEEFF'</code>
///   <para>Creates a rectangle with a custom fill color.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficePowerPointShape")]
public sealed class AddOfficePowerPointShapeCommand : PSCmdlet
{
    /// <summary>Target slide that will receive the shape.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true)]
    public PowerPointSlide Slide { get; set; } = null!;

    /// <summary>Shape geometry preset name (e.g., Rectangle, Ellipse, Line).</summary>
    [Parameter]
    public string ShapeType { get; set; } = "Rectangle";

    /// <summary>Left offset (in points) from the slide origin.</summary>
    [Parameter]
    public double X { get; set; } = 50;

    /// <summary>Top offset (in points) from the slide origin.</summary>
    [Parameter]
    public double Y { get; set; } = 50;

    /// <summary>Shape width in points.</summary>
    [Parameter]
    public double Width { get; set; } = 200;

    /// <summary>Shape height in points.</summary>
    [Parameter]
    public double Height { get; set; } = 100;

    /// <summary>Optional name assigned to the shape.</summary>
    [Parameter]
    public string? Name { get; set; }

    /// <summary>Fill color (hex or named color).</summary>
    [Parameter]
    public string? FillColor { get; set; }

    /// <summary>Outline color (hex or named color).</summary>
    [Parameter]
    public string? OutlineColor { get; set; }

    /// <summary>Outline width in points.</summary>
    [Parameter]
    public double? OutlineWidth { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        try
        {
            if (Width <= 0)
            {
                throw new ArgumentOutOfRangeException(nameof(Width), "Width must be greater than 0.");
            }

            if (Height <= 0)
            {
                throw new ArgumentOutOfRangeException(nameof(Height), "Height must be greater than 0.");
            }

            if (OutlineWidth is < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(OutlineWidth), "OutlineWidth cannot be negative.");
            }

            var shapeType = ResolveShapeType(ShapeType);
            var shape = Slide.AddShapePoints(shapeType, X, Y, Width, Height, Name);

            var fill = NormalizeColor(FillColor);
            if (fill != null)
            {
                shape.FillColor = fill;
            }

            var outline = NormalizeColor(OutlineColor);
            if (outline != null)
            {
                shape.OutlineColor = outline;
            }

            if (OutlineWidth.HasValue)
            {
                shape.OutlineWidthPoints = OutlineWidth.Value;
            }

            WriteObject(shape);
        }
        catch (Exception ex)
        {
            WriteError(new ErrorRecord(ex, "PowerPointAddShapeFailed", ErrorCategory.InvalidOperation, Slide));
        }
    }

    private static string? NormalizeColor(string? color)
    {
        if (string.IsNullOrWhiteSpace(color))
        {
            return null;
        }

        var parsed = Color.Parse(color);
        var hex = parsed.ToHex().ToLowerInvariant();
        return hex.Length > 6 ? hex.Substring(0, 6) : hex;
    }

    private static ShapeTypeValues ResolveShapeType(string? shapeType)
    {
        if (string.IsNullOrWhiteSpace(shapeType))
        {
            return ShapeTypeValues.Rectangle;
        }

        var property = typeof(ShapeTypeValues).GetProperty(
            shapeType,
            BindingFlags.Public | BindingFlags.Static | BindingFlags.IgnoreCase);

        if (property == null)
        {
            throw new PSArgumentException($"Unknown shape type '{shapeType}'.", nameof(ShapeType));
        }

        return (ShapeTypeValues)property.GetValue(null)!;
    }
}
