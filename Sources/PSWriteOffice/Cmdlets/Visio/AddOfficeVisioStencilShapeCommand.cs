using System.Management.Automation;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Stencils;
using PSWriteOffice.Services.Visio;

namespace PSWriteOffice.Cmdlets.Visio;

/// <summary>Adds a stencil shape to the current Visio page.</summary>
[Cmdlet(VerbsCommon.Add, "OfficeVisioStencilShape", DefaultParameterSetName = CatalogNameParameterSet)]
[Alias("VisioStencil")]
[OutputType(typeof(VisioShape))]
public sealed class AddOfficeVisioStencilShapeCommand : PSCmdlet
{
    private const string CatalogObjectParameterSet = "CatalogObject";
    private const string CatalogNameParameterSet = "CatalogName";
    private const string BuiltInParameterSet = "BuiltIn";

    /// <summary>Target page. Optional inside <c>VisioPage</c> or <c>New-OfficeVisio</c>.</summary>
    [Parameter(ValueFromPipeline = true)]
    public VisioPage? Page { get; set; }

    /// <summary>Catalog object containing the stencil shape.</summary>
    [Parameter(Mandatory = true, ParameterSetName = CatalogObjectParameterSet)]
    public VisioStencilCatalog? CatalogObject { get; set; }

    /// <summary>Catalog previously registered in the active Visio DSL scope.</summary>
    [Parameter(ParameterSetName = CatalogNameParameterSet)]
    [Alias("CatalogName")]
    public string? Catalog { get; set; }

    /// <summary>Built-in OfficeIMO stencil catalog containing the shape.</summary>
    [Parameter(ParameterSetName = BuiltInParameterSet)]
    public OfficeVisioBuiltInStencilCatalog BuiltIn { get; set; } = OfficeVisioBuiltInStencilCatalog.All;

    /// <summary>Stencil id, name, master name, keyword, alias, or tag.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    [Alias("Name")]
    public string Stencil { get; set; } = string.Empty;

    /// <summary>DSL key used by connector commands.</summary>
    [Parameter]
    public string? Key { get; set; }

    /// <summary>X coordinate of the stencil shape center.</summary>
    [Parameter]
    public double X { get; set; } = 1;

    /// <summary>Y coordinate of the stencil shape center.</summary>
    [Parameter]
    public double Y { get; set; } = 1;

    /// <summary>Optional shape width. Omit to use the stencil default width.</summary>
    [Parameter]
    public double? Width { get; set; }

    /// <summary>Optional shape height. Omit to use the stencil default height.</summary>
    [Parameter]
    public double? Height { get; set; }

    /// <summary>Text placed inside the shape. Omit to use the stencil display name.</summary>
    [Parameter(Position = 1)]
    public string? Text { get; set; }

    /// <summary>Optional shape name.</summary>
    [Parameter]
    public string? ShapeName { get; set; }

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
        var catalog = ResolveCatalog(context);
        var id = string.IsNullOrWhiteSpace(Key) ? Stencil : Key!;
        var shape = Width.HasValue || Height.HasValue
            ? page.AddStencilShape(catalog, Stencil, id, X, Y, Width ?? catalog.Get(Stencil).DefaultWidth, Height ?? catalog.Get(Stencil).DefaultHeight, Text)
            : page.AddStencilShape(catalog, Stencil, id, X, Y, Text);

        VisioShapeCommandUtilities.ApplyShapeStyle(shape, ShapeName ?? Key, NameU, FillColor, LineColor, LineWeight, LinePattern, FillPattern, Angle);
        context?.RegisterShape(Key, shape);
        WriteObject(shape);
    }

    private VisioStencilCatalog ResolveCatalog(VisioDslContext? context)
    {
        if (ParameterSetName == CatalogObjectParameterSet)
        {
            return CatalogObject!;
        }

        if (ParameterSetName == BuiltInParameterSet)
        {
            return VisioStencilCommandUtilities.GetBuiltInCatalog(BuiltIn);
        }

        if (context != null)
        {
            return context.ResolveStencilCatalog(Catalog);
        }

        return string.IsNullOrWhiteSpace(Catalog)
            ? VisioStencils.All
            : throw new PSInvalidOperationException("A named stencil catalog can only be resolved inside New-OfficeVisio.");
    }
}
