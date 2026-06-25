#pragma warning disable CS1591
using System.Globalization;
using System.IO;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Imports normalized CSV/TSV text into an Excel workbook through OfficeIMO.</summary>
/// <example>
///   <summary>Import a semicolon-delimited regional export as a typed worksheet table.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$result = Import-OfficeExcelDelimitedText -Path .\Report.xlsx `
///     -SourcePath .\sales-pl.csv `
///     -Delimiter ';' `
///     -CultureName 'pl-PL' `
///     -SheetName Sales `
///     -PassThru
/// $result | Format-List SheetName,Range,RowCount,ColumnCount,Delimiter</code>
///   <para>Normalizes delimited text through OfficeIMO, performs culture-aware value conversion, and writes the result as an Excel table unless -NoTable is used.</para>
/// </example>
[Cmdlet(VerbsData.Import, "OfficeExcelDelimitedText", DefaultParameterSetName = ParameterSetPath, SupportsShouldProcess = true, ConfirmImpact = ConfirmImpact.Medium)]
[Alias("ExcelDelimitedImport", "ExcelCsvImport")]
[OutputType(typeof(PSObject))]
public sealed class ImportOfficeExcelDelimitedTextCommand : PSCmdlet
{
    private const string ParameterSetPath = "Path";
    private const string ParameterSetDocument = "Document";

    /// <summary>Workbook path.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("Path", "FilePath")]
    public string InputPath { get; set; } = string.Empty;

    /// <summary>Workbook document.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public ExcelDocument Document { get; set; } = null!;

    /// <summary>Delimited text source path.</summary>
    [Parameter(Mandatory = true)]
    public string SourcePath { get; set; } = string.Empty;

    /// <summary>Delimiter character. When omitted, it is detected.</summary>
    [Parameter]
    public char? Delimiter { get; set; }
    /// <summary>Worksheet name to create or inspect.</summary>
    [Parameter]
    public string? SheetName { get; set; }
    /// <summary>Culture name for number and date conversion.</summary>
    [Parameter]
    public string? CultureName { get; set; }
    /// <summary>Treat the first row as data.</summary>
    [Parameter]
    public SwitchParameter NoHeader { get; set; }
    /// <summary>Import rows without creating an Excel table.</summary>
    [Parameter]
    public SwitchParameter NoTable { get; set; }
    /// <summary>Keep imported values as text.</summary>
    [Parameter]
    public SwitchParameter NoTypeConversion { get; set; }
    /// <summary>Emit a result object.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    protected override void ProcessRecord()
    {
        var source = SessionState.Path.GetUnresolvedProviderPathFromPSPath(SourcePath);
        if (!File.Exists(source))
        {
            throw new FileNotFoundException($"Delimited text file '{source}' was not found.", source);
        }

        var target = ResolveTargetPath();
        if (!ShouldProcess(target ?? "Excel document", "Import delimited text into Excel workbook"))
        {
            return;
        }

        using var workbook = ResolveWorkbook(target);
        var result = workbook.Document.ImportDelimitedFile(source, new ExcelDelimitedImportOptions
        {
            Delimiter = Delimiter,
            SheetName = SheetName,
            HeadersInFirstRow = !NoHeader.IsPresent,
            CreateTable = !NoTable.IsPresent,
            ConvertNumbersAndDates = !NoTypeConversion.IsPresent,
            Culture = string.IsNullOrWhiteSpace(CultureName) ? CultureInfo.InvariantCulture : CultureInfo.GetCultureInfo(CultureName!)
        });

        workbook.SaveIfOwned();
        if (PassThru.IsPresent)
        {
            var output = new PSObject();
            output.Properties.Add(new PSNoteProperty("SheetName", result.ImportResult.SheetName));
            output.Properties.Add(new PSNoteProperty("TableName", result.ImportResult.TableName));
            output.Properties.Add(new PSNoteProperty("Range", result.ImportResult.Range));
            output.Properties.Add(new PSNoteProperty("RowCount", result.ImportResult.RowCount));
            output.Properties.Add(new PSNoteProperty("ColumnCount", result.ImportResult.ColumnCount));
            output.Properties.Add(new PSNoteProperty("Delimiter", result.Delimiter.ToString()));
            output.Properties.Add(new PSNoteProperty("Warnings", result.Warnings));
            WriteObject(output);
        }
    }

    private string? ResolveTargetPath()
    {
        if (!string.Equals(ParameterSetName, ParameterSetPath, System.StringComparison.OrdinalIgnoreCase))
        {
            return Document.FilePath;
        }

        return SessionState.Path.GetUnresolvedProviderPathFromPSPath(InputPath);
    }

    private ExcelWorkbookCommandScope ResolveWorkbook(string? targetPath)
    {
        if (!string.Equals(ParameterSetName, ParameterSetPath, System.StringComparison.OrdinalIgnoreCase))
        {
            return new ExcelWorkbookCommandScope(Document, ownsDocument: false);
        }

        if (string.IsNullOrWhiteSpace(targetPath))
        {
            throw new PSArgumentException("Specify a workbook path.", nameof(InputPath));
        }

        var resolvedPath = targetPath!;
        var directory = Path.GetDirectoryName(resolvedPath);
        if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
        {
            Directory.CreateDirectory(directory);
        }

        var document = File.Exists(resolvedPath)
            ? ExcelDocumentService.LoadDocument(resolvedPath, readOnly: false, autoSave: false)
            : ExcelDocumentService.CreateDocument(resolvedPath, autoSave: false);

        return new ExcelWorkbookCommandScope(document, ownsDocument: true);
    }
}
#pragma warning restore CS1591
