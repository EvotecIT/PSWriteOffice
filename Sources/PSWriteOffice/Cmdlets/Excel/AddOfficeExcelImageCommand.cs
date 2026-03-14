using System;
using System.IO;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Adds an image anchored to a worksheet cell.</summary>
/// <example>
///   <summary>Insert an image from disk at B2.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ExcelSheet 'Data' { Add-OfficeExcelImage -Address 'B2' -Path .\logo.png -WidthPixels 120 -HeightPixels 40 }</code>
///   <para>Anchors the image to cell B2.</para>
/// </example>
/// <example>
///   <summary>Insert an image from a URL.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ExcelSheet 'Data' { Add-OfficeExcelImage -Row 1 -Column 1 -Url 'https://example.org/logo.png' }</code>
///   <para>Downloads and anchors the image to cell A1.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeExcelImage", DefaultParameterSetName = ParameterSetContextPath)]
[Alias("ExcelImage")]
public sealed class AddOfficeExcelImageCommand : PSCmdlet
{
    private const string ParameterSetContextPath = "ContextPath";
    private const string ParameterSetContextUrl = "ContextUrl";
    private const string ParameterSetDocumentPath = "DocumentPath";
    private const string ParameterSetDocumentUrl = "DocumentUrl";

    /// <summary>Workbook to operate on outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocumentPath)]
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocumentUrl)]
    public ExcelDocument Document { get; set; } = null!;

    /// <summary>Worksheet name when using <see cref="Document"/>.</summary>
    [Parameter(ParameterSetName = ParameterSetDocumentPath)]
    [Parameter(ParameterSetName = ParameterSetDocumentUrl)]
    public string? Sheet { get; set; }

    /// <summary>Worksheet index (0-based) when using <see cref="Document"/>.</summary>
    [Parameter(ParameterSetName = ParameterSetDocumentPath)]
    [Parameter(ParameterSetName = ParameterSetDocumentUrl)]
    public int? SheetIndex { get; set; }

    /// <summary>Image file path.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetContextPath)]
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetDocumentPath)]
    public string Path { get; set; } = string.Empty;

    /// <summary>Image URL to download.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetContextUrl)]
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetDocumentUrl)]
    public string Url { get; set; } = string.Empty;

    /// <summary>1-based row index.</summary>
    [Parameter]
    public int? Row { get; set; }

    /// <summary>1-based column index.</summary>
    [Parameter]
    public int? Column { get; set; }

    /// <summary>A1-style cell address (e.g., A1, C5).</summary>
    [Parameter]
    public string? Address { get; set; }

    /// <summary>Image width in pixels.</summary>
    [Parameter]
    public int WidthPixels { get; set; } = 96;

    /// <summary>Image height in pixels.</summary>
    [Parameter]
    public int HeightPixels { get; set; } = 32;

    /// <summary>Horizontal offset in pixels from the cell origin.</summary>
    [Parameter]
    public int OffsetXPixels { get; set; }

    /// <summary>Vertical offset in pixels from the cell origin.</summary>
    [Parameter]
    public int OffsetYPixels { get; set; }

    /// <summary>Emit the worksheet after inserting the image.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (WidthPixels <= 0 || HeightPixels <= 0)
        {
            throw new PSArgumentException("WidthPixels and HeightPixels must be greater than zero.");
        }

        var sheet = ResolveSheet();
        var (row, column) = ExcelHostExtensions.ResolveCellAddress(Row, Column, Address);

        if (ParameterSetName == ParameterSetContextPath || ParameterSetName == ParameterSetDocumentPath)
        {
            var resolved = SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path);
            if (!File.Exists(resolved))
            {
                throw new FileNotFoundException($"Image file '{resolved}' was not found.", resolved);
            }

            var bytes = File.ReadAllBytes(resolved);
            sheet.AddImageAt(row, column, bytes, GetContentType(resolved), WidthPixels, HeightPixels, OffsetXPixels, OffsetYPixels);
        }
        else
        {
            sheet.AddImageFromUrlAt(row, column, Url, WidthPixels, HeightPixels, OffsetXPixels, OffsetYPixels);
        }

        if (PassThru.IsPresent)
        {
            WriteObject(sheet);
        }
    }

    private ExcelSheet ResolveSheet()
    {
        if (ParameterSetName == ParameterSetDocumentPath || ParameterSetName == ParameterSetDocumentUrl)
        {
            if (Document == null)
            {
                throw new PSArgumentException("Provide an Excel document.");
            }

            return ExcelSheetResolver.Resolve(Document, Sheet, SheetIndex);
        }

        var context = ExcelDslContext.Require(this);
        return context.RequireSheet();
    }

    private static string GetContentType(string path)
    {
        var ext = System.IO.Path.GetExtension(path)?.ToLowerInvariant();
        return ext switch
        {
            ".jpg" or ".jpeg" => "image/jpeg",
            ".gif" => "image/gif",
            ".bmp" => "image/bmp",
            ".tif" or ".tiff" => "image/tiff",
            _ => "image/png"
        };
    }
}
