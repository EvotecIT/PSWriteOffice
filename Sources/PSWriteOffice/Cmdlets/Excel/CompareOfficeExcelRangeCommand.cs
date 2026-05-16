using System;
using System.IO;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Compares two Excel worksheets or ranges and returns cell-level differences.</summary>
/// <example>
///   <summary>Compare two sheets in the same workbook.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Compare-OfficeExcelRange -Path .\Report.xlsx -LeftSheet Current -RightSheet Expected</code>
///   <para>Compares the used ranges of the two worksheets.</para>
/// </example>
[Cmdlet(VerbsData.Compare, "OfficeExcelRange", DefaultParameterSetName = ParameterSetPath)]
[Alias("Compare-OfficeExcelSheet", "ExcelCompare")]
[OutputType(typeof(ExcelRangeDifference))]
public sealed class CompareOfficeExcelRangeCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";
    private const string ParameterSetPath = "Path";

    /// <summary>Left workbook path.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("Path", "FilePath", "LeftPath")]
    public string InputPath { get; set; } = string.Empty;

    /// <summary>Left workbook object.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public ExcelDocument Document { get; set; } = null!;

    /// <summary>Optional right workbook path. Defaults to the left workbook.</summary>
    [Parameter(ParameterSetName = ParameterSetPath)]
    public string? RightPath { get; set; }

    /// <summary>Optional right workbook object. Defaults to the left workbook.</summary>
    [Parameter(ParameterSetName = ParameterSetContext)]
    [Parameter(ParameterSetName = ParameterSetDocument)]
    public ExcelDocument? RightDocument { get; set; }

    /// <summary>Left worksheet name. Defaults to the current sheet in a sheet block, otherwise the first sheet.</summary>
    [Parameter]
    public string? LeftSheet { get; set; }

    /// <summary>Left worksheet index.</summary>
    [Parameter]
    public int? LeftSheetIndex { get; set; }

    /// <summary>Right worksheet name. Defaults to the left sheet.</summary>
    [Parameter]
    public string? RightSheet { get; set; }

    /// <summary>Right worksheet index.</summary>
    [Parameter]
    public int? RightSheetIndex { get; set; }

    /// <summary>Left A1 range. Defaults to the left worksheet used range.</summary>
    [Parameter]
    public string? LeftRange { get; set; }

    /// <summary>Right A1 range. Defaults to the right worksheet used range.</summary>
    [Parameter]
    public string? RightRange { get; set; }

    /// <summary>Compare strings after trimming whitespace.</summary>
    [Parameter]
    public SwitchParameter TrimStrings { get; set; }

    /// <summary>Compare strings case-insensitively.</summary>
    [Parameter]
    public SwitchParameter IgnoreCase { get; set; }

    /// <summary>Treat null and empty string values as different.</summary>
    [Parameter]
    public SwitchParameter StrictNullEmpty { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        ExcelDocument? leftDocument = null;
        ExcelDocument? openedLeft = null;
        ExcelDocument? openedRight = null;

        try
        {
            leftDocument = ResolveLeftDocument(out openedLeft);
            var rightDocument = ResolveRightDocument(leftDocument, out openedRight);
            var leftSheet = ResolveLeftSheet(leftDocument);
            var rightSheet = ResolveRightSheet(rightDocument, leftSheet.Name);

            var options = new ExcelRangeCompareOptions
            {
                TrimStrings = TrimStrings.IsPresent,
                IgnoreCase = IgnoreCase.IsPresent,
                TreatNullAndEmptyStringAsEqual = !StrictNullEmpty.IsPresent
            };

            var differences = leftDocument.CompareRanges(
                leftSheet,
                string.IsNullOrWhiteSpace(LeftRange) ? leftSheet.GetUsedRangeA1() : LeftRange!,
                rightSheet,
                string.IsNullOrWhiteSpace(RightRange) ? rightSheet.GetUsedRangeA1() : RightRange!,
                options);

            WriteObject(differences, enumerateCollection: true);
        }
        finally
        {
            openedRight?.Dispose();
            openedLeft?.Dispose();
        }
    }

    private ExcelDocument ResolveLeftDocument(out ExcelDocument? opened)
    {
        opened = null;
        if (ParameterSetName == ParameterSetPath)
        {
            var resolvedPath = SessionState.Path.GetUnresolvedProviderPathFromPSPath(InputPath);
            opened = ExcelDocumentService.LoadDocument(resolvedPath, readOnly: true, autoSave: false);
            return opened;
        }

        return ParameterSetName == ParameterSetDocument
            ? Document ?? throw new PSArgumentException("Provide an Excel document.")
            : ExcelDslContext.Require(this).Document;
    }

    private ExcelDocument ResolveRightDocument(ExcelDocument leftDocument, out ExcelDocument? opened)
    {
        opened = null;
        if (ParameterSetName == ParameterSetPath && !string.IsNullOrWhiteSpace(RightPath))
        {
            var resolvedPath = SessionState.Path.GetUnresolvedProviderPathFromPSPath(RightPath!);
            opened = ExcelDocumentService.LoadDocument(resolvedPath, readOnly: true, autoSave: false);
            return opened;
        }

        return RightDocument ?? leftDocument;
    }

    private ExcelSheet ResolveLeftSheet(ExcelDocument document)
    {
        if (ParameterSetName == ParameterSetContext && string.IsNullOrWhiteSpace(LeftSheet) && !LeftSheetIndex.HasValue)
        {
            return ExcelDslContext.Require(this).RequireSheet();
        }

        return ExcelSheetResolver.Resolve(document, LeftSheet, LeftSheetIndex);
    }

    private ExcelSheet ResolveRightSheet(ExcelDocument document, string leftSheetName)
    {
        if (!string.IsNullOrWhiteSpace(RightSheet) || RightSheetIndex.HasValue)
        {
            return ExcelSheetResolver.Resolve(document, RightSheet, RightSheetIndex);
        }

        return document[leftSheetName];
    }
}
