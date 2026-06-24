using System.Management.Automation;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Adds workbook-level slicer cache metadata.</summary>
/// <example>
///   <summary>Add slicer cache metadata for a pivot field.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$slicer = Add-OfficeExcelSlicer -Path .\Report.xlsx -Name RegionSlicer -SourceName Region -PivotTableName SalesPivot -PassThru
/// Get-OfficeExcelDataModel -Path .\Report.xlsx |
///     Select-Object -ExpandProperty SlicerCacheCount</code>
///   <para>Writes slicer cache package metadata through OfficeIMO. Excel may still be required to materialize full slicer UI shapes.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeExcelSlicer", DefaultParameterSetName = ParameterSetContext)]
[Alias("ExcelSlicer")]
[OutputType(typeof(PSObject))]
public sealed class AddOfficeExcelSlicerCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";
    private const string ParameterSetPath = "Path";

    /// <summary>Workbook path to update.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("Path", "FilePath")]
    public string InputPath { get; set; } = string.Empty;

    /// <summary>Workbook to update outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public ExcelDocument Document { get; set; } = null!;

    /// <summary>Slicer cache name.</summary>
    [Parameter(Mandatory = true)]
    public string Name { get; set; } = string.Empty;

    /// <summary>Source field, table column, or pivot field name.</summary>
    [Parameter]
    public string? SourceName { get; set; }

    /// <summary>Pivot table name the slicer is intended to filter.</summary>
    [Parameter]
    public string? PivotTableName { get; set; }

    /// <summary>Caller-supplied slicer cache XML. When provided, OfficeIMO writes it as-is.</summary>
    [Parameter]
    public string? Xml { get; set; }

    /// <summary>Emit metadata about the added package part.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        using var workbook = ExcelWorkbookCommandService.ResolveWorkbook(this, ParameterSetName, InputPath, Document, readOnly: false);
        ExtendedPart part = workbook.Document.AddWorkbookSlicerCache(new ExcelSlicerCacheOptions
        {
            Name = Name,
            SourceName = SourceName,
            PivotTableName = PivotTableName,
            Xml = Xml
        });

        string contentType = part.ContentType;
        workbook.SaveIfOwned();

        if (PassThru.IsPresent)
        {
            var result = new PSObject();
            result.Properties.Add(new PSNoteProperty("Name", Name));
            result.Properties.Add(new PSNoteProperty("Kind", "Slicer"));
            result.Properties.Add(new PSNoteProperty("ContentType", contentType));
            WriteObject(result);
        }
    }
}
