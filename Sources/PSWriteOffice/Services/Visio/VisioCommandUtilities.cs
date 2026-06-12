using System;
using System.IO;
using System.Management.Automation;
using OfficeIMO.Drawing;
using OfficeIMO.Visio;

namespace PSWriteOffice.Services.Visio;

internal static class VisioCommandUtilities
{
    internal static string ResolvePath(PSCmdlet cmdlet, string path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            throw new PSArgumentException("Path cannot be empty.", nameof(path));
        }

        var providerPath = cmdlet.SessionState.Path.GetUnresolvedProviderPathFromPSPath(path);
        return Path.IsPathRooted(providerPath)
            ? providerPath
            : Path.Combine(cmdlet.SessionState.Path.CurrentFileSystemLocation.Path, providerPath);
    }

    internal static void EnsureDirectory(string path)
    {
        var directory = Path.GetDirectoryName(path);
        if (!string.IsNullOrWhiteSpace(directory) && !Directory.Exists(directory))
        {
            Directory.CreateDirectory(directory);
        }
    }

    internal static VisioDocument ResolveDocument(PSCmdlet cmdlet, VisioDocument? document, string? path)
    {
        if (document != null)
        {
            return document;
        }

        if (string.IsNullOrWhiteSpace(path))
        {
            throw new PSArgumentException("Provide -Document or -Path.", nameof(path));
        }

        return VisioDocument.Load(ResolvePath(cmdlet, path!));
    }

    internal static VisioSvgSaveOptions BuildSvgOptions(
        int pageIndex,
        double? pixelsPerInch,
        string? backgroundColor,
        bool transparent,
        bool noText,
        bool noStencilArtwork,
        bool noConnectorLabels,
        bool noConnectorLabelOverlapResolution,
        bool includeXmlDeclaration)
    {
        var options = new VisioSvgSaveOptions
        {
            PageIndex = pageIndex,
            RenderText = !noText,
            RenderStencilArtwork = !noStencilArtwork,
            RenderConnectorLabels = !noConnectorLabels,
            ResolveConnectorLabelOverlaps = !noConnectorLabelOverlapResolution,
            IncludeXmlDeclaration = includeXmlDeclaration
        };

        if (pixelsPerInch.HasValue)
        {
            options.PixelsPerInch = pixelsPerInch.Value;
        }

        options.BackgroundColor = ResolveBackgroundColor(backgroundColor, transparent);
        return options;
    }

    internal static VisioPngSaveOptions BuildPngOptions(
        PSCmdlet cmdlet,
        int pageIndex,
        double? pixelsPerInch,
        string? backgroundColor,
        bool transparent,
        bool noText,
        string? fontFilePath,
        string? fontFaceName,
        int? fontCollectionIndex,
        bool noStencilArtwork,
        bool noConnectorLabels,
        bool noConnectorLabelOverlapResolution,
        int? supersampling)
    {
        var options = new VisioPngSaveOptions
        {
            PageIndex = pageIndex,
            RenderText = !noText,
            FontFilePath = string.IsNullOrWhiteSpace(fontFilePath) ? null : ResolvePath(cmdlet, fontFilePath!),
            FontFaceName = fontFaceName,
            FontCollectionIndex = fontCollectionIndex,
            RenderStencilArtwork = !noStencilArtwork,
            RenderConnectorLabels = !noConnectorLabels,
            ResolveConnectorLabelOverlaps = !noConnectorLabelOverlapResolution
        };

        if (pixelsPerInch.HasValue)
        {
            options.PixelsPerInch = pixelsPerInch.Value;
        }

        if (supersampling.HasValue)
        {
            options.Supersampling = supersampling.Value;
        }

        options.BackgroundColor = ResolveBackgroundColor(backgroundColor, transparent);
        return options;
    }

    private static OfficeColor? ResolveBackgroundColor(string? color, bool transparent)
    {
        if (transparent)
        {
            return null;
        }

        return string.IsNullOrWhiteSpace(color) ? OfficeColor.White : OfficeColor.Parse(color!);
    }
}
