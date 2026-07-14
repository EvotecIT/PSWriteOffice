using System;
using System.IO;
using System.Management.Automation;
using OfficeIMO.Excel.OpenDocument;
using OfficeIMO.OpenDocument;
using OfficeIMO.PowerPoint.OpenDocument;
using OfficeIMO.Word.OpenDocument;

namespace PSWriteOffice.Cmdlets.OpenDocument;

/// <summary>Converts native ODT, ODS, or ODP content to Word, Excel, or PowerPoint with fidelity evidence.</summary>
[Cmdlet(VerbsData.ConvertFrom, "OfficeOpenDocument", SupportsShouldProcess = true)]
public sealed class ConvertFromOfficeOpenDocumentCommand : PSCmdlet
{
    /// <summary>Path to an ODT, ODS, or ODP file.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Path { get; set; } = string.Empty;

    /// <summary>Destination DOCX, XLSX, or PPTX path.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    public string OutputPath { get; set; } = string.Empty;

    /// <summary>Optional ODT-to-Word conversion settings.</summary>
    [Parameter]
    public WordOpenDocumentConversionOptions? WordOptions { get; set; }

    /// <summary>Optional ODS-to-Excel conversion settings.</summary>
    [Parameter]
    public ExcelOpenDocumentConversionOptions? ExcelOptions { get; set; }

    /// <summary>Optional ODP-to-PowerPoint conversion settings.</summary>
    [Parameter]
    public PowerPointOpenDocumentConversionOptions? PowerPointOptions { get; set; }

    /// <summary>Throw when the conversion approximates, skips, or cannot map a feature.</summary>
    [Parameter]
    public SwitchParameter FailOnLoss { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var input = SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path);
        var output = SessionState.Path.GetUnresolvedProviderPathFromPSPath(OutputPath);
        var source = OdfDocument.Load(input);
        ValidateOutputExtension(output, source.Kind);
        if (!ShouldProcess(output, "Convert OpenDocument to Office document")) return;
        Directory.CreateDirectory(System.IO.Path.GetDirectoryName(output) ?? SessionState.Path.CurrentFileSystemLocation.Path);
        switch (source)
        {
            case OdtDocument text:
                var wordResult = text.ToWordDocumentResult(WordOptions);
                if (FailOnLoss.IsPresent) wordResult.RequireNoLoss();
                wordResult.Value.Save(output);
                WriteObject(wordResult);
                break;
            case OdsDocument spreadsheet:
                var excelResult = spreadsheet.ToExcelDocumentResult(ExcelOptions);
                if (FailOnLoss.IsPresent) excelResult.RequireNoLoss();
                excelResult.Value.Save(output);
                WriteObject(excelResult);
                break;
            case OdpPresentation presentation:
                var presentationResult = presentation.ToPowerPointPresentationResult(PowerPointOptions);
                if (FailOnLoss.IsPresent) presentationResult.RequireNoLoss();
                presentationResult.Value.Save(output);
                WriteObject(presentationResult);
                break;
            default:
                throw new InvalidOperationException("Unsupported OpenDocument kind.");
        }
    }

    private static void ValidateOutputExtension(string outputPath, OdfDocumentKind kind)
    {
        var expected = kind switch
        {
            OdfDocumentKind.Text => ".docx",
            OdfDocumentKind.Spreadsheet => ".xlsx",
            OdfDocumentKind.Presentation => ".pptx",
            _ => throw new InvalidOperationException("Unsupported OpenDocument kind.")
        };
        var actual = System.IO.Path.GetExtension(outputPath);
        if (!string.Equals(actual, expected, StringComparison.OrdinalIgnoreCase))
        {
            throw new PSArgumentException($"OutputPath must use the {expected} extension for {kind} content.", nameof(OutputPath));
        }
    }
}
