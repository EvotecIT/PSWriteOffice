using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Splits a PDF into page, range, count, or bookmark files.</summary>
/// <example>
///   <summary>Split a PDF into page files.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$pages = Split-OfficePdf -Path .\Examples\Documents\Combined.pdf -OutputDirectory .\Examples\Documents\Pages -Prefix 'combined-page'
/// $pages | Select-Object Name, Length</code>
///   <para>Creates one output PDF for each page and returns the written files.</para>
/// </example>
[Cmdlet(VerbsCommon.Split, "OfficePdf")]
[OutputType(typeof(FileInfo))]
public sealed class SplitOfficePdfCommand : PSCmdlet
{
    /// <summary>Input PDF path.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Output directory.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    public string OutputDirectory { get; set; } = string.Empty;

    /// <summary>Output file prefix.</summary>
    [Parameter]
    public string Prefix { get; set; } = "page";

    /// <summary>Create one PDF for each consecutive group with this many pages.</summary>
    [Parameter]
    public int PagesPerDocument { get; set; }

    /// <summary>Create one PDF for each supplied page range or selection, such as 1-3 or 1,3.</summary>
    [Parameter]
    [Alias("Range", "PageRanges")]
    public string[]? PageRange { get; set; }

    /// <summary>Create one PDF for each supplied bookmark title.</summary>
    [Parameter]
    [Alias("Bookmark", "BookmarkTitle")]
    public string[]? BookmarkName { get; set; }

    /// <summary>Create one PDF for every readable bookmark when -BookmarkName is not supplied.</summary>
    [Parameter]
    public SwitchParameter ByBookmark { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var outputDirectory = PdfCommandUtilities.ResolvePath(this, OutputDirectory);
        Directory.CreateDirectory(outputDirectory);
        var document = PdfDocument.Open(PdfCommandUtilities.ResolvePath(this, Path));
        var outputs = CreateOutputs(document);
        foreach (var output in outputs)
        {
            var outputPath = PdfCommandUtilities.GetUniquePath(outputDirectory, Prefix + "-" + output.Name + ".pdf");
            output.Document.Save(outputPath);
            WriteObject(new FileInfo(outputPath));
        }
    }

    private IReadOnlyList<PdfSplitOutput> CreateOutputs(PdfDocument document)
    {
        var modes = 0;
        if (PagesPerDocument > 0)
        {
            modes++;
        }

        if (PageRange != null && PageRange.Length > 0)
        {
            modes++;
        }

        if (ByBookmark.IsPresent || BookmarkName is { Length: > 0 })
        {
            modes++;
        }

        if (modes > 1)
        {
            throw new PSArgumentException("Use only one split mode: -PagesPerDocument, -PageRange, or -ByBookmark/-BookmarkName.");
        }

        if (PagesPerDocument < 0)
        {
            throw new PSArgumentException("-PagesPerDocument must be greater than zero.", nameof(PagesPerDocument));
        }

        if (PagesPerDocument > 0)
        {
            var ranges = GetPageGroups(document.Inspect().PageCount, PagesPerDocument);
            var documents = document.Pages.Split(ranges);
            return CreateRangeOutputs(documents, ranges);
        }

        if (PageRange != null && PageRange.Length > 0)
        {
            var selections = PageRange.Select(PdfPageSelection.Parse).ToArray();
            var documents = document.Pages.Split(selections);
            return documents
                .Select((pdf, index) => new PdfSplitOutput(pdf, PdfCommandUtilities.GetSafeFileName(PageRange[index])))
                .ToArray();
        }

        if (ByBookmark.IsPresent || BookmarkName is { Length: > 0 })
        {
            var names = BookmarkName ?? Array.Empty<string>();
            var ranges = document.Pages.BookmarkPageRanges(names);
            var documents = document.Pages.Split(ranges.Select(range => range.PageRange));
            return documents
                .Select((pdf, index) => new PdfSplitOutput(pdf, PdfCommandUtilities.GetSafeFileName(ranges[index].Title)))
                .ToArray();
        }

        var split = document.Pages.Split();
        return split
            .Select((pdf, index) => new PdfSplitOutput(pdf, (index + 1).ToString(CultureInfo.InvariantCulture)))
            .ToArray();
    }

    private static IReadOnlyList<PdfSplitOutput> CreateRangeOutputs(IReadOnlyList<PdfDocument> documents, IReadOnlyList<PdfPageRange> ranges)
    {
        return documents
            .Select((pdf, index) => new PdfSplitOutput(pdf, ranges[index].ToString()))
            .ToArray();
    }

    private static IReadOnlyList<PdfPageRange> GetPageGroups(int pageCount, int pagesPerDocument)
    {
        if (pageCount <= 0)
        {
            throw new PSInvalidOperationException("PDF does not contain any readable pages.");
        }

        var ranges = new List<PdfPageRange>();
        for (var firstPage = 1; firstPage <= pageCount; firstPage += pagesPerDocument)
        {
            var lastPage = Math.Min(firstPage + pagesPerDocument - 1, pageCount);
            ranges.Add(PdfPageRange.From(firstPage, lastPage));
        }

        return ranges;
    }

    private sealed class PdfSplitOutput
    {
        internal PdfSplitOutput(PdfDocument document, string name)
        {
            Document = document;
            Name = name;
        }

        internal PdfDocument Document { get; }

        internal string Name { get; }
    }
}
