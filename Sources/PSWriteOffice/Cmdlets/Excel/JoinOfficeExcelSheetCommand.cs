using System;
using System.IO;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Appends or merges rows from one worksheet into another.</summary>
/// <example>
///   <summary>Append source rows below a target sheet.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Join-OfficeExcelSheet -Path .\Report.xlsx -TargetSheet Combined -SourceSheet Data</code>
///   <para>Copies rows from Data into Combined, skipping the source header row by default.</para>
/// </example>
[Cmdlet(VerbsCommon.Join, "OfficeExcelSheet", DefaultParameterSetName = ParameterSetContext)]
[Alias("Merge-OfficeExcelSheet", "ExcelSheetJoin", "ExcelSheetMerge")]
[OutputType(typeof(ExcelWorksheetMergeResult))]
public sealed class JoinOfficeExcelSheetCommand : PSCmdlet
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

    /// <summary>Target worksheet name. Defaults to the current sheet inside an ExcelSheet block.</summary>
    [Parameter]
    public string? TargetSheet { get; set; }

    /// <summary>Target worksheet index when using a workbook object or path.</summary>
    [Parameter]
    public int? TargetSheetIndex { get; set; }

    /// <summary>Source worksheet name.</summary>
    [Parameter(Mandatory = true)]
    public string SourceSheet { get; set; } = string.Empty;

    /// <summary>Optional source workbook object for cross-workbook joins.</summary>
    [Parameter]
    public ExcelDocument? SourceDocument { get; set; }

    /// <summary>Optional source workbook path for cross-workbook joins.</summary>
    [Parameter]
    public string? SourcePath { get; set; }

    /// <summary>Source A1 range to copy. Defaults to the source used range.</summary>
    [Parameter]
    public string? SourceRange { get; set; }

    /// <summary>1-based target start row. Defaults to appending after the target used range.</summary>
    [Parameter]
    public int? TargetStartRow { get; set; }

    /// <summary>1-based target start column. Defaults to the source range start column.</summary>
    [Parameter]
    public int? TargetStartColumn { get; set; }

    /// <summary>Treat the first source row as data instead of a header row.</summary>
    [Parameter]
    public SwitchParameter NoSourceHeader { get; set; }

    /// <summary>Include the source header row in copied rows.</summary>
    [Parameter]
    public SwitchParameter IncludeSourceHeader { get; set; }

    /// <summary>Match source columns to target columns by header text.</summary>
    [Parameter]
    public SwitchParameter MatchColumnsByHeader { get; set; }

    /// <summary>1-based target header row when matching columns by header.</summary>
    [Parameter]
    public int? TargetHeaderRow { get; set; }

    /// <summary>Blank rows to leave before appended data.</summary>
    [Parameter]
    public int BlankRowsBefore { get; set; }

    /// <summary>Allow copied values to replace existing target cells.</summary>
    [Parameter]
    public SwitchParameter OverwriteExistingCells { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        ExcelDocument? targetDocument = null;
        ExcelDocument? openedSource = null;
        var disposeTarget = false;

        try
        {
            ValidateSourceOptions();
            targetDocument = ResolveTargetDocument(out disposeTarget);
            var targetSheet = ResolveTargetSheet(targetDocument);
            var sourceDocument = ResolveSourceDocument(targetDocument, out openedSource);
            var sourceSheet = sourceDocument[SourceSheet];

            var result = targetDocument.JoinWorksheets(targetSheet, sourceSheet, BuildOptions());
            if (disposeTarget)
            {
                targetDocument.Save(false);
            }

            WriteObject(result);
        }
        finally
        {
            openedSource?.Dispose();
            if (disposeTarget)
            {
                targetDocument?.Dispose();
            }
        }
    }

    private ExcelWorksheetMergeOptions BuildOptions()
    {
        return new ExcelWorksheetMergeOptions
        {
            SourceRange = SourceRange,
            TargetStartRow = TargetStartRow,
            TargetStartColumn = TargetStartColumn,
            SourceHasHeader = !NoSourceHeader.IsPresent,
            IncludeSourceHeader = IncludeSourceHeader.IsPresent,
            MatchColumnsByHeader = MatchColumnsByHeader.IsPresent,
            TargetHeaderRow = TargetHeaderRow,
            BlankRowsBefore = BlankRowsBefore,
            OverwriteExistingCells = OverwriteExistingCells.IsPresent
        };
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

        return ParameterSetName == ParameterSetDocument
            ? Document ?? throw new PSArgumentException("Provide an Excel document.")
            : ExcelDslContext.Require(this).Document;
    }

    private ExcelSheet ResolveTargetSheet(ExcelDocument document)
    {
        if (ParameterSetName == ParameterSetContext && string.IsNullOrWhiteSpace(TargetSheet) && !TargetSheetIndex.HasValue)
        {
            return ExcelDslContext.Require(this).RequireSheet();
        }

        return ExcelSheetResolver.Resolve(document, TargetSheet, TargetSheetIndex);
    }

    private ExcelDocument ResolveSourceDocument(ExcelDocument targetDocument, out ExcelDocument? openedSource)
    {
        openedSource = null;
        if (!string.IsNullOrWhiteSpace(SourcePath))
        {
            var resolvedSource = SessionState.Path.GetUnresolvedProviderPathFromPSPath(SourcePath!);
            openedSource = ExcelDocumentService.LoadDocument(resolvedSource, readOnly: true, autoSave: false);
            return openedSource;
        }

        return SourceDocument ?? targetDocument;
    }

    private void ValidateSourceOptions()
    {
        if (SourceDocument != null && !string.IsNullOrWhiteSpace(SourcePath))
        {
            throw new PSArgumentException("Specify either -SourceDocument or -SourcePath, not both.");
        }
    }
}
