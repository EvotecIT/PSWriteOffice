using System;
using System.Data;
using System.IO;
using System.Management.Automation;
using System.Threading.Tasks;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Reads the used range from an Excel workbook.</summary>
/// <para>Returns rows as PSCustomObjects by default, with optional hashtable or DataTable output for scripting and interoperability.</para>
/// <example>
///   <summary>Read the used range and produce a quick status count.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$rows = Get-OfficeExcelUsedRange -Path .\report.xlsx -Sheet Data
/// $rows |
///     Group-Object -Property Status |
///     Select-Object -Property Name, Count</code>
///   <para>Reads the sheet's used range, treats the first row as headers, and summarizes a status column.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeExcelUsedRange", DefaultParameterSetName = ParameterSetPath)]
[OutputType(typeof(PSObject), typeof(System.Collections.Hashtable), typeof(DataTable))]
public sealed class GetOfficeExcelUsedRangeCommand : AsyncPSCmdlet
{
    private const string ParameterSetPath = "Path";
    private const string ParameterSetUri = "Uri";
    private const string ParameterSetDocument = "Document";

    /// <summary>Path to the workbook.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("FilePath", "Path")]
    public string InputPath { get; set; } = string.Empty;

    /// <summary>Remote workbook URI to read.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetUri)]
    [Alias("Url")]
    public Uri? Uri { get; set; }

    /// <summary>Workbook to inspect.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public ExcelDocument Document { get; set; } = null!;

    /// <summary>Allow HTTP workbook downloads in addition to HTTPS.</summary>
    [Parameter(ParameterSetName = ParameterSetUri)]
    public SwitchParameter AllowHttp { get; set; }

    /// <summary>Worksheet name to read; defaults to the first sheet.</summary>
    [Parameter]
    public string? Sheet { get; set; }

    /// <summary>Zero-based worksheet index to read; defaults to the first sheet.</summary>
    [Parameter]
    public int? SheetIndex { get; set; }

    /// <summary>Use the first row as column headers.</summary>
    [Parameter]
    public bool HeadersInFirstRow { get; set; } = true;

    /// <summary>Prefer decimals instead of doubles for numeric values.</summary>
    [Parameter]
    public SwitchParameter NumericAsDecimal { get; set; }

    /// <summary>Emit rows as hashtables instead of PSCustomObjects.</summary>
    [Parameter]
    public SwitchParameter AsHashtable { get; set; }

    /// <summary>Emit the raw DataTable instead of row objects.</summary>
    [Parameter]
    public SwitchParameter AsDataTable { get; set; }

    /// <inheritdoc />
    protected override async Task ProcessRecordAsync()
    {
        var options = ExcelReadOutputService.CreateOptions(NumericAsDecimal.IsPresent);
        using var reader = await CreateReaderAsync(options).ConfigureAwait(false);
        var sheetReader = ExcelReadOutputService.ResolveSheetReader(reader, Sheet, SheetIndex);
        var usedRange = sheetReader.GetUsedRangeA1();
        var table = sheetReader.ReadRangeAsDataTable(usedRange, HeadersInFirstRow);

        ExcelReadOutputService.WriteOutput(this, table, AsDataTable.IsPresent, AsHashtable.IsPresent);
    }

    private async Task<ExcelDocumentReader> CreateReaderAsync(ExcelReadOptions options)
    {
        if (ParameterSetName == ParameterSetDocument)
        {
            return Document.CreateReader(options);
        }

        if (ParameterSetName == ParameterSetUri)
        {
            if (Uri == null)
            {
                throw new PSArgumentException("Workbook URI was not provided.", nameof(Uri));
            }

            return await ExcelDocumentReader.OpenAsync(Uri, options, ExcelHttpLoadService.CreateOptions(AllowHttp), CancelToken)
                .ConfigureAwait(false);
        }

        var resolvedPath = SessionState.Path.GetUnresolvedProviderPathFromPSPath(InputPath);
        if (!File.Exists(resolvedPath))
        {
            throw new FileNotFoundException($"File '{resolvedPath}' was not found.", resolvedPath);
        }

        return ExcelDocumentReader.Open(resolvedPath, options);
    }
}
