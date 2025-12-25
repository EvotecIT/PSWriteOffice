using System;
using System.Collections.Concurrent;
using System.Diagnostics;
using System.IO;
using OfficeIMO.PowerPoint;

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

        var presentation = PowerPointPresentation.Create(filePath);
        Presentations[presentation] = filePath;
        return presentation;
    }

    /// <summary>Loads an existing presentation.</summary>
    public static PowerPointPresentation LoadPresentation(string filePath)
    {
        if (!File.Exists(filePath))
        {
            throw new FileNotFoundException($"File {filePath} doesn't exist.", filePath);
        }

        var presentation = PowerPointPresentation.Open(filePath);
        Presentations[presentation] = filePath;
        return presentation;
    }

    /// <summary>Saves and optionally opens the presentation.</summary>
    public static void SavePresentation(PowerPointPresentation presentation, bool show)
    {
        if (!Presentations.TryGetValue(presentation, out var filePath))
        {
            throw new ArgumentException("Presentation was not created or loaded via this service.", nameof(presentation));
        }

        presentation.Save();

        if (show)
        {
            var startInfo = new ProcessStartInfo
            {
                FileName = filePath,
                UseShellExecute = true
            };
            Process.Start(startInfo);
        }

        presentation.Dispose();
        Presentations.TryRemove(presentation, out _);
    }
}
