using System;
using System.IO;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Copies a worksheet within a workbook or from another workbook.</summary>
/// <example>
///   <summary>Copy a worksheet in the current workbook.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Copy-OfficeExcelSheet -Path .\Report.xlsx -SourceSheet Data -NewName DataCopy</code>
///   <para>Creates a copy of the Data worksheet.</para>
/// </example>
[Cmdlet(VerbsCommon.Copy, "OfficeExcelSheet", DefaultParameterSetName = ParameterSetContext)]
[Alias("ExcelSheetCopy")]
[OutputType(typeof(ExcelSheet))]
public sealed class CopyOfficeExcelSheetCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";
    private const string ParameterSetPath = "Path";

    /// <summary>Target workbook path to update.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("Path", "FilePath")]
    public string InputPath { get; set; } = string.Empty;

    /// <summary>Target workbook to update outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public ExcelDocument Document { get; set; } = null!;

    /// <summary>Optional source workbook object for cross-workbook copies.</summary>
    [Parameter(ParameterSetName = ParameterSetContext)]
    [Parameter(ParameterSetName = ParameterSetDocument)]
    [Parameter(ParameterSetName = ParameterSetPath)]
    public ExcelDocument? SourceDocument { get; set; }

    /// <summary>Optional source workbook path for cross-workbook copies.</summary>
    [Parameter(ParameterSetName = ParameterSetContext)]
    [Parameter(ParameterSetName = ParameterSetDocument)]
    [Parameter(ParameterSetName = ParameterSetPath)]
    public string? SourcePath { get; set; }

    /// <summary>Worksheet to copy. Defaults to the current sheet inside an ExcelSheet block.</summary>
    [Parameter(Position = 1)]
    [Alias("Sheet", "WorksheetName")]
    public string? SourceSheet { get; set; }

    /// <summary>Name for the copied worksheet.</summary>
    [Parameter(Mandatory = true, Position = 2)]
    [Alias("Name", "DestinationSheet")]
    public string NewName { get; set; } = string.Empty;

    /// <summary>Controls how invalid destination sheet names are handled.</summary>
    [Parameter]
    public SheetNameValidationMode ValidationMode { get; set; } = SheetNameValidationMode.Sanitize;

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        ExcelDocument? target = null;
        ExcelDocument? openedSource = null;
        var disposeTarget = false;

        try
        {
            target = ResolveTargetDocument(out disposeTarget);
            var sourceSheet = ResolveSourceSheetName(target);
            ValidateSourceOptions();

            ExcelSheet copied;
            if (!string.IsNullOrWhiteSpace(SourcePath))
            {
                var resolvedSource = SessionState.Path.GetUnresolvedProviderPathFromPSPath(SourcePath!);
                openedSource = ExcelDocumentService.LoadDocument(resolvedSource, readOnly: true, autoSave: false);
                copied = target.CopyWorksheetFrom(openedSource, sourceSheet, NewName, ValidationMode);
            }
            else if (SourceDocument != null)
            {
                copied = target.CopyWorksheetFrom(SourceDocument, sourceSheet, NewName, ValidationMode);
            }
            else
            {
                copied = target.CopyWorksheet(sourceSheet, NewName, ValidationMode);
            }

            if (disposeTarget)
            {
                target.Save(false);
            }

            WriteObject(copied);
        }
        finally
        {
            openedSource?.Dispose();
            if (disposeTarget)
            {
                target?.Dispose();
            }
        }
    }

    private ExcelDocument ResolveTargetDocument(out bool dispose)
    {
        dispose = false;
        if (ParameterSetName == ParameterSetPath)
        {
            var resolvedPath = SessionState.Path.GetUnresolvedProviderPathFromPSPath(InputPath);
            if (!File.Exists(resolvedPath))
            {
                throw new FileNotFoundException($"File '{resolvedPath}' was not found.", resolvedPath);
            }

            dispose = true;
            return ExcelDocumentService.LoadDocument(resolvedPath, readOnly: false, autoSave: false);
        }

        if (ParameterSetName == ParameterSetDocument)
        {
            return Document ?? throw new PSArgumentException("Provide an Excel document.");
        }

        return ExcelDslContext.Require(this).Document;
    }

    private string ResolveSourceSheetName(ExcelDocument target)
    {
        if (!string.IsNullOrWhiteSpace(SourceSheet))
        {
            return SourceSheet!;
        }

        if (ParameterSetName == ParameterSetContext)
        {
            return ExcelDslContext.Require(this).RequireSheet().Name;
        }

        if (target.Sheets.Count == 0)
        {
            throw new InvalidOperationException("Workbook contains no worksheets.");
        }

        return target.Sheets[0].Name;
    }

    private void ValidateSourceOptions()
    {
        if (SourceDocument != null && !string.IsNullOrWhiteSpace(SourcePath))
        {
            throw new PSArgumentException("Specify either -SourceDocument or -SourcePath, not both.");
        }
    }
}
