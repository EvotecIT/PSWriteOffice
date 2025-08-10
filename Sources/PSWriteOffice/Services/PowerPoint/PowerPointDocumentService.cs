using System;
using System.Collections.Concurrent;
using System.Diagnostics;
using System.IO;
using ShapeCrawler;

namespace PSWriteOffice.Services.PowerPoint;

public static class PowerPointDocumentService
{
    private static readonly ConcurrentDictionary<Presentation, (string Path, bool IsNew)> Presentations = new();

    public static Presentation CreatePresentation(string filePath)
    {
        var presentation = new Presentation();
        // ShapeCrawler always creates a presentation with 1 default slide
        // and doesn't allow removing the last slide
        Presentations[presentation] = (filePath, true);
        return presentation;
    }

    public static Presentation LoadPresentation(string filePath)
    {
        if (!File.Exists(filePath))
        {
            throw new FileNotFoundException($"File {filePath} doesn't exist.", filePath);
        }

        var presentation = new Presentation(filePath);
        Presentations[presentation] = (filePath, false);
        return presentation;
    }

    public static void SavePresentation(Presentation presentation, bool show)
    {
        if (!Presentations.TryGetValue(presentation, out var info))
        {
            throw new ArgumentException("Presentation was not created or loaded via this service.", nameof(presentation));
        }

        if (info.IsNew)
        {
            presentation.Save(info.Path);
            Presentations[presentation] = (info.Path, false);
        }
        else
        {
            presentation.Save();
        }

        if (show)
        {
            var startInfo = new ProcessStartInfo
            {
                FileName = info.Path,
                UseShellExecute = true
            };
            Process.Start(startInfo);
        }

        presentation.Dispose();
        Presentations.TryRemove(presentation, out _);
    }
}
