using System.Management.Automation;
using OfficeIMO.Visio;
using PSWriteOffice.Services.Visio;

namespace PSWriteOffice.Cmdlets.Visio;

/// <summary>Adds a rectangle shape to the current Visio page.</summary>
[Cmdlet(VerbsCommon.Add, "OfficeVisioRectangle")]
[Alias("VisioRectangle", "VisioRect")]
[OutputType(typeof(VisioShape))]
public sealed class AddOfficeVisioRectangleCommand : PSCmdlet
{
    /// <summary>Target page. Optional inside <c>VisioPage</c> or <c>New-OfficeVisio</c>.</summary>
    [Parameter(ValueFromPipeline = true)]
    public VisioPage? Page { get; set; }

    /// <summary>DSL key used by connector commands.</summary>
    [Parameter]
    public string? Key { get; set; }

    /// <summary>X coordinate of the shape origin.</summary>
    [Parameter]
    public double X { get; set; } = 1;

    /// <summary>Y coordinate of the shape origin.</summary>
    [Parameter]
    public double Y { get; set; } = 1;

    /// <summary>Shape width.</summary>
    [Parameter]
    public double Width { get; set; } = 2;

    /// <summary>Shape height.</summary>
    [Parameter]
    public double Height { get; set; } = 1;

    /// <summary>Text placed inside the shape.</summary>
    [Parameter(Position = 0)]
    public string? Text { get; set; }

    /// <summary>Measurement unit for coordinates and dimensions.</summary>
    [Parameter]
    public VisioMeasurementUnit Unit { get; set; } = VisioMeasurementUnit.Inches;

    /// <summary>Optional shape name.</summary>
    [Parameter]
    public string? Name { get; set; }

    /// <summary>Optional universal shape name.</summary>
    [Parameter]
    public string? NameU { get; set; }

    /// <summary>Fill color name or hex value.</summary>
    [Parameter]
    public string? FillColor { get; set; }

    /// <summary>Line color name or hex value.</summary>
    [Parameter]
    public string? LineColor { get; set; }

    /// <summary>Line weight.</summary>
    [Parameter]
    public double? LineWeight { get; set; }

    /// <summary>Line pattern.</summary>
    [Parameter]
    public int? LinePattern { get; set; }

    /// <summary>Fill pattern.</summary>
    [Parameter]
    public int? FillPattern { get; set; }

    /// <summary>Shape angle in radians.</summary>
    [Parameter]
    public double? Angle { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var context = VisioDslContext.Current;
        var page = Page ?? VisioDslContext.Require(this).RequirePage();
        var shape = page.AddRectangle(X, Y, Width, Height, Text, Unit);
        VisioShapeCommandUtilities.ApplyShapeStyle(shape, Name ?? Key, NameU, FillColor, LineColor, LineWeight, LinePattern, FillPattern, Angle);
        context?.RegisterShape(page, Key, shape);
        WriteObject(shape);
    }
}
