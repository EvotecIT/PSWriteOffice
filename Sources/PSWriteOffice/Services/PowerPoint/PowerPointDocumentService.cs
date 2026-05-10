using System;
using System.Collections.Concurrent;
using System.IO;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Services;

namespace PSWriteOffice.Services.PowerPoint;

/// <summary>Helper methods bridging DSL cmdlets with OfficeIMO PowerPoint presentations.</summary>
public static class PowerPointDocumentService
{
    private static readonly ConcurrentDictionary<PowerPointPresentation, string> Presentations = new();

    /// <summary>Creates a new presentation at the specified path.</summary>
    public static PowerPointPresentation CreatePresentation(string filePath)
    {
        if (string.IsNullOrWhiteSpace(filePath))
        {
            throw new ArgumentException("File path cannot be empty.", nameof(filePath));
        }

        var resolvedPath = Path.GetFullPath(filePath);
        var presentation = PowerPointPresentation.Create(resolvedPath);
        Presentations[presentation] = resolvedPath;
        return presentation;
    }

    /// <summary>Loads an existing presentation.</summary>
    public static PowerPointPresentation LoadPresentation(string filePath)
    {
        var resolvedPath = Path.GetFullPath(filePath);
        if (!File.Exists(resolvedPath))
        {
            throw new FileNotFoundException($"File {resolvedPath} doesn't exist.", resolvedPath);
        }

        var presentation = PowerPointPresentation.Open(resolvedPath);
        Presentations[presentation] = resolvedPath;
        return presentation;
    }

    /// <summary>Saves and optionally opens the presentation.</summary>
    public static void SavePresentation(PowerPointPresentation presentation, bool show)
    {
        if (!Presentations.TryGetValue(presentation, out var filePath))
        {
            throw new ArgumentException("Presentation was not created or loaded via this service.", nameof(presentation));
        }

        var resolvedPath = Path.GetFullPath(filePath);
        presentation.Save();
        presentation.Dispose();
        Presentations.TryRemove(presentation, out _);

        if (show)
        {
            FileOpenService.Open(resolvedPath);
        }
    }
}
