using System;
using System.IO;
using System.Management.Automation;
using ChartForgeX.VisualArtifacts;
using OfficeIMO.ChartForgeX;

namespace PSWriteOffice.Services.Visuals;

/// <summary>Provides shared, typed ChartForgeX-to-OfficeIMO conversion parameters.</summary>
public abstract class OfficeVisualCommandBase : PSCmdlet
{
    /// <summary>SVG fidelity policy used by OfficeIMO.ChartForgeX.</summary>
    [Parameter]
    public OfficeVisualSvgPolicy SvgPolicy { get; set; } = OfficeVisualSvgPolicy.PreserveVector;

    /// <summary>Optional output width in Office points.</summary>
    [Parameter]
    public double? Width { get; set; }

    /// <summary>Optional output height in Office points.</summary>
    [Parameter]
    public double? Height { get; set; }

    /// <summary>Conversion factor from ChartForgeX pixels to Office points.</summary>
    [Parameter]
    public double PointsPerPixel { get; set; } = 0.75D;

    /// <summary>Optional OfficeIMO SVG import element limit.</summary>
    [Parameter]
    public int? MaximumSvgElements { get; set; }

    /// <summary>Optional stable identifier for an SVG file input.</summary>
    [Parameter]
    public string? Id { get; set; }

    /// <summary>Optional title for an SVG file input.</summary>
    [Parameter]
    public string? Title { get; set; }

    /// <summary>Optional accessible description for an SVG file input.</summary>
    [Parameter]
    public string? AlternativeText { get; set; }

    /// <summary>Converts an artifact or reuses an existing conversion result.</summary>
    protected OfficeVisualConversionResult ResolveVisual(object inputObject)
    {
        object value = inputObject is PSObject psObject ? psObject.BaseObject : inputObject;
        if (value is OfficeVisualConversionResult converted)
        {
            RejectConversionOverrides();
            return converted;
        }
        if (value is OfficeVisualSource source)
        {
            RejectSourceMetadataOverrides();
            return source.ToOfficeVisual(CreateOptions());
        }

        if (value is FileInfo fileInfo)
        {
            return ConvertSvgFile(fileInfo.FullName);
        }
        if (value is string path)
        {
            return ConvertSvgFile(SessionState.Path.GetUnresolvedProviderPathFromPSPath(path));
        }

        if (value is not VisualArtifact artifact)
        {
            throw new PSArgumentException(
                "InputObject must be a ChartForgeX VisualArtifact, OfficeVisualSource, OfficeVisualConversionResult, or SVG file path.",
                nameof(inputObject));
        }

        return artifact.ToOfficeVisual(CreateOptions());
    }

    private void RejectConversionOverrides()
    {
        string[] names = { nameof(SvgPolicy), nameof(Width), nameof(Height), nameof(PointsPerPixel), nameof(MaximumSvgElements), nameof(Id), nameof(Title), nameof(AlternativeText) };
        foreach (string name in names)
        {
            if (MyInvocation.BoundParameters.ContainsKey(name))
            {
                throw new PSArgumentException($"-{name} cannot be used with an existing OfficeVisualConversionResult.", name);
            }
        }
    }

    private void RejectSourceMetadataOverrides()
    {
        foreach (string name in new[] { nameof(Id), nameof(Title), nameof(AlternativeText) })
        {
            if (MyInvocation.BoundParameters.ContainsKey(name))
            {
                throw new PSArgumentException($"-{name} cannot be used with an existing OfficeVisualSource.", name);
            }
        }
    }

    private OfficeVisualConversionResult ConvertSvgFile(string path)
    {
        string fullPath = Path.GetFullPath(path);
        if (!File.Exists(fullPath))
        {
            throw new FileNotFoundException($"SVG visual '{fullPath}' was not found.", fullPath);
        }
        if (!string.Equals(Path.GetExtension(fullPath), ".svg", StringComparison.OrdinalIgnoreCase))
        {
            throw new PSArgumentException("Portable Office visual input must be an .svg file.", nameof(path));
        }
        var source = new OfficeVisualSource(File.ReadAllBytes(fullPath))
        {
            Id = string.IsNullOrWhiteSpace(Id) ? Path.GetFileNameWithoutExtension(fullPath) : Id!,
            Title = Title ?? string.Empty,
            AlternativeText = AlternativeText ?? string.Empty
        };
        return source.ToOfficeVisual(CreateOptions());
    }

    private OfficeVisualConversionOptions CreateOptions()
    {
        var options = new OfficeVisualConversionOptions
        {
            SvgPolicy = SvgPolicy,
            PointsPerPixel = PointsPerPixel,
            WidthPoints = Width,
            HeightPoints = Height
        };
        if (MaximumSvgElements.HasValue)
        {
            options.MaximumSvgElements = MaximumSvgElements.Value;
        }
        return options;
    }
}
