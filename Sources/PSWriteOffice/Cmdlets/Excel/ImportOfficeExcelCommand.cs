using System;
using System.IO;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Imports rows from an Excel workbook as PowerShell objects.</summary>
/// <para>Provides an ImportExcel-style read command over the OfficeIMO reader pipeline.</para>
/// <example>
///   <summary>Import worksheet rows and filter pending items.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$rows = Import-OfficeExcel -Path .\Report.xlsx -WorksheetName Data -NumericAsDecimal
/// $rows |
///     Where-Object Status -eq 'Pending' |
///     Export-Csv -Path .\PendingRows.csv -NoTypeInformation</code>
///   <para>Reads the used range on the Data worksheet, emits PSCustomObjects, and filters them in PowerShell.</para>
/// </example>
[Cmdlet(VerbsData.Import, "OfficeExcel", DefaultParameterSetName = ParameterSetPath)]
[Alias("ExcelImport")]
public sealed class ImportOfficeExcelCommand : PSCmdlet
{
    private const string ParameterSetPath = "Path";
    private const string ParameterSetUri = "Uri";
    private const string ParameterSetDocument = "Document";

    /// <summary>Workbook path to import.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipelineByPropertyName = true, ParameterSetName = ParameterSetPath)]
    [Alias("FilePath", "InputPath", "FullName")]
    public string? Path { get; set; }

    /// <summary>Remote workbook URI to import.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipelineByPropertyName = true, ParameterSetName = ParameterSetUri)]
    [Alias("Url")]
    public Uri? Uri { get; set; }

    /// <summary>Workbook document to import from.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public ExcelDocument? Document { get; set; }

    /// <summary>Allow HTTP workbook downloads in addition to HTTPS.</summary>
    [Parameter(ParameterSetName = ParameterSetUri)]
    public SwitchParameter AllowHttp { get; set; }

    /// <summary>Worksheet name to read; defaults to the first sheet.</summary>
    [Parameter(ValueFromPipelineByPropertyName = true)]
    [Alias("Sheet")]
    public string? WorksheetName { get; set; }

    /// <summary>Zero-based worksheet index to read.</summary>
    [Parameter(ValueFromPipelineByPropertyName = true)]
    public int? SheetIndex { get; set; }

    /// <summary>Optional A1 range to read. When omitted, the used range is imported.</summary>
    [Parameter(ValueFromPipelineByPropertyName = true)]
    public string? Range { get; set; }

    /// <summary>Starting row for an explicit rectangular range.</summary>
    [Parameter]
    public int? StartRow { get; set; }

    /// <summary>Ending row for an explicit rectangular range.</summary>
    [Parameter]
    public int? EndRow { get; set; }

    /// <summary>Starting column for an explicit rectangular range.</summary>
    [Parameter]
    public int? StartColumn { get; set; }

    /// <summary>Ending column for an explicit rectangular range.</summary>
    [Parameter]
    public int? EndColumn { get; set; }

    /// <summary>Treat all rows as data and generate column names instead of using the first row as headers.</summary>
    [Parameter]
    public SwitchParameter NoHeader { get; set; }

    /// <summary>Prefer decimals instead of doubles for numeric values.</summary>
    [Parameter]
    public SwitchParameter NumericAsDecimal { get; set; }

    /// <summary>Emit rows as hashtables instead of PSCustomObjects.</summary>
    [Parameter]
    public SwitchParameter AsHashtable { get; set; }

    /// <summary>Emit a DataTable instead of enumerating row objects.</summary>
    [Parameter]
    public SwitchParameter AsDataTable { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (!string.IsNullOrWhiteSpace(Range) && HasCoordinateRange())
        {
            throw new PSArgumentException("Specify either -Range or coordinate bounds, but not both.");
        }

        var options = ExcelReadOutputService.CreateOptions(NumericAsDecimal.IsPresent);
        using var reader = CreateReader(options);
        var sheet = ExcelReadOutputService.ResolveSheetReader(reader, WorksheetName, SheetIndex);
        var range = ResolveRange(sheet);
        var table = sheet.ReadRangeAsDataTable(range, headersInFirstRow: !NoHeader.IsPresent);

        ExcelReadOutputService.WriteOutput(this, table, AsDataTable.IsPresent, AsHashtable.IsPresent);
    }

    private ExcelDocumentReader CreateReader(ExcelReadOptions options)
    {
        if (ParameterSetName == ParameterSetDocument)
        {
            if (Document == null)
            {
                throw new PSArgumentException("Excel document was not provided.", nameof(Document));
            }

            return ExcelDocumentReader.Wrap(Document._spreadSheetDocument, options);
        }

        if (ParameterSetName == ParameterSetUri)
        {
            if (Uri == null)
            {
                throw new PSArgumentException("Workbook URI was not provided.", nameof(Uri));
            }

            return ExcelDocumentReader.Open(Uri, options, ExcelHttpLoadService.CreateOptions(AllowHttp));
        }

        if (string.IsNullOrWhiteSpace(Path))
        {
            throw new PSArgumentException("Workbook path was not provided.", nameof(Path));
        }

        var resolvedPath = SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path!);
        if (!File.Exists(resolvedPath))
        {
            throw new FileNotFoundException($"File '{resolvedPath}' was not found.", resolvedPath);
        }

        return ExcelDocumentReader.Open(resolvedPath, options);
    }

    private string ResolveRange(ExcelSheetReader sheet)
    {
        if (!string.IsNullOrWhiteSpace(Range))
        {
            return Range!;
        }

        if (HasCoordinateRange())
        {
            if (!StartRow.HasValue || !EndRow.HasValue || !StartColumn.HasValue || !EndColumn.HasValue)
            {
                throw new PSArgumentException("StartRow, EndRow, StartColumn, and EndColumn must all be provided when using coordinate bounds.");
            }

            if (StartRow.Value < 1 || EndRow.Value < 1 || StartColumn.Value < 1 || EndColumn.Value < 1)
            {
                throw new PSArgumentException("Coordinate bounds must be 1 or greater.");
            }

            if (StartRow.Value > EndRow.Value)
            {
                throw new PSArgumentException("StartRow must be less than or equal to EndRow.");
            }

            if (StartColumn.Value > EndColumn.Value)
            {
                throw new PSArgumentException("StartColumn must be less than or equal to EndColumn.");
            }

            return $"{A1.CellReference(StartRow.Value, StartColumn.Value)}:{A1.CellReference(EndRow.Value, EndColumn.Value)}";
        }

        return sheet.GetUsedRangeA1();
    }

    private bool HasCoordinateRange()
    {
        return StartRow.HasValue || EndRow.HasValue || StartColumn.HasValue || EndColumn.HasValue;
    }
}
