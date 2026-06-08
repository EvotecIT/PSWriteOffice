using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Copies a worksheet within a workbook or from another workbook.</summary>
/// <example>
///   <summary>Copy a worksheet and confirm it exists.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$proof = @(
///     Copy-OfficeExcelSheet -Path .\Report.xlsx -SourceSheet Data -NewName DataCopy
///     Get-OfficeExcelSummary -Path .\Report.xlsx -IncludeSheets |
///         Select-Object -ExpandProperty Sheets |
///         Where-Object Name -eq 'DataCopy'
/// )
/// $proof</code>
///   <para>Creates a copy of the Data worksheet and reads the workbook summary to verify the new tab.</para>
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
        using var targetWorkbook = ExcelWorkbookCommandService.ResolveWorkbook(this, ParameterSetName, InputPath, Document, readOnly: false);
        var target = targetWorkbook.Document;
        var sourceSheet = ExcelWorkbookCommandService.ResolveSheetNameOrCurrent(this, target, ParameterSetName, SourceSheet);
        using var sourceWorkbook = ExcelWorkbookCommandService.ResolveSourceWorkbook(this, target, SourceDocument, SourcePath, readOnly: true);

        var copied = sourceWorkbook.Document == target
            ? target.CopyWorksheet(sourceSheet, NewName, ValidationMode)
            : target.CopyWorksheetFrom(sourceWorkbook.Document, sourceSheet, NewName, ValidationMode);

        targetWorkbook.SaveIfOwned();
        WriteObject(copied);
    }
}
