using System;
using System.Management.Automation;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Adds explicit workbook package metadata such as connection or query-table XML.</summary>
/// <example>
///   <summary>Add connection metadata and query-table metadata to an existing workbook.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Add-OfficeExcelPackageMetadata -Path .\Report.xlsx -Kind Connection -Xml $connectionsXml
/// Add-OfficeExcelPackageMetadata -Path .\Report.xlsx -Kind QueryTable -WorksheetName Data -Xml $queryTableXml</code>
///   <para>Adds caller-supplied XML metadata parts. OfficeIMO preserves these parts but does not refresh external queries.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeExcelPackageMetadata", DefaultParameterSetName = ParameterSetContext)]
[Alias("ExcelPackageMetadata", "ExcelConnectionMetadata")]
[OutputType(typeof(PSObject))]
public sealed class AddOfficeExcelPackageMetadataCommand : PSCmdlet
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

    /// <summary>Metadata kind to add.</summary>
    [Parameter(Mandatory = true)]
    [ValidateSet("Connection", "QueryTable")]
    public string Kind { get; set; } = "Connection";

    /// <summary>XML payload to add as package metadata.</summary>
    [Parameter(Mandatory = true)]
    public string Xml { get; set; } = string.Empty;

    /// <summary>Worksheet for query-table metadata. Defaults to the current DSL sheet, or the first worksheet outside the DSL.</summary>
    [Parameter]
    [Alias("Sheet", "SheetName", "Worksheet")]
    public string? WorksheetName { get; set; }

    /// <summary>Emit metadata about the added package part.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        using var workbook = ExcelWorkbookCommandService.ResolveWorkbook(this, ParameterSetName, InputPath, Document, readOnly: false);
        var document = workbook.Document;

        string? sheetName = null;
        ExtendedPart part;
        if (string.Equals(Kind, "QueryTable", StringComparison.OrdinalIgnoreCase))
        {
            sheetName = ExcelWorkbookCommandService.ResolveSheetNameOrCurrent(this, document, ParameterSetName, WorksheetName);
            part = document.AddWorksheetQueryTableMetadata(sheetName, Xml);
        }
        else
        {
            part = document.AddWorkbookConnectionMetadata(Xml);
        }

        string contentType = part.ContentType;
        workbook.SaveIfOwned();

        if (PassThru.IsPresent)
        {
            var result = new PSObject();
            result.Properties.Add(new PSNoteProperty("Kind", Kind));
            result.Properties.Add(new PSNoteProperty("WorksheetName", sheetName));
            result.Properties.Add(new PSNoteProperty("ContentType", contentType));
            WriteObject(result);
        }
    }
}
