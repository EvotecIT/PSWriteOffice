using System;
using System.Data;
using System.IO;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Reads an explicit A1 range from an Excel workbook.</summary>
/// <para>Returns rows as PSCustomObjects by default, with optional hashtable or DataTable output for scripting and interoperability.</para>
/// <example>
///   <summary>Read a rectangular range as objects.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Get-OfficeExcelRange -Path .\report.xlsx -Sheet 'Data' -Range 'A1:C10'</code>
///   <para>Uses the first row as headers and returns each remaining row as a PSCustomObject.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeExcelRange", DefaultParameterSetName = ParameterSetPath)]
[OutputType(typeof(PSObject), typeof(System.Collections.Hashtable), typeof(DataTable))]
public sealed class GetOfficeExcelRangeCommand : PSCmdlet
{
    private const string ParameterSetPath = "Path";
    private const string ParameterSetDocument = "Document";

    /// <summary>Path to the workbook.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("FilePath", "Path")]
    public string InputPath { get; set; } = string.Empty;

    /// <summary>Workbook to inspect.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public ExcelDocument Document { get; set; } = null!;

    /// <summary>A1 range to read.</summary>
    [Parameter(Mandatory = true)]
    public string Range { get; set; } = string.Empty;

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
    protected override void ProcessRecord()
    {
        var options = ExcelReadOutputService.CreateOptions(NumericAsDecimal.IsPresent);
        using var reader = CreateReader(options);
        var sheetReader = ExcelReadOutputService.ResolveSheetReader(reader, Sheet, SheetIndex);
        var table = sheetReader.ReadRangeAsDataTable(Range, HeadersInFirstRow);

        ExcelReadOutputService.WriteOutput(this, table, AsDataTable.IsPresent, AsHashtable.IsPresent);
    }

    private ExcelDocumentReader CreateReader(ExcelReadOptions options)
    {
        if (ParameterSetName == ParameterSetDocument)
        {
            return ExcelDocumentReader.Wrap(Document._spreadSheetDocument, options);
        }

        var resolvedPath = SessionState.Path.GetUnresolvedProviderPathFromPSPath(InputPath);
        if (!File.Exists(resolvedPath))
        {
            throw new FileNotFoundException($"File '{resolvedPath}' was not found.", resolvedPath);
        }

        return ExcelDocumentReader.Open(resolvedPath, options);
    }
}
