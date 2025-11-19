using System.IO;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Opens an existing Excel workbook.</summary>
/// <para>Returns the underlying <see cref="ExcelDocument"/> so callers can inspect or reuse it in DSL pipelines.</para>
/// <example>
///   <summary>Load a workbook in read-only mode.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$workbook = Get-OfficeExcel -Path .\report.xlsx -ReadOnly</code>
///   <para>Loads <c>report.xlsx</c> for inspection without enabling writes.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeExcel")]
public sealed class GetOfficeExcelCommand : PSCmdlet
{
    /// <summary>Path to the workbook to load.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    [Alias("Path", "FilePath")]
    public string InputPath { get; set; } = string.Empty;

    /// <summary>Open the file in read-only mode.</summary>
    [Parameter]
    public SwitchParameter ReadOnly { get; set; }

    /// <summary>Enable automatic saves on the underlying document.</summary>
    [Parameter]
    public SwitchParameter AutoSave { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var resolvedPath = SessionState.Path.GetUnresolvedProviderPathFromPSPath(InputPath);
        if (!File.Exists(resolvedPath))
        {
            throw new FileNotFoundException($"File '{resolvedPath}' was not found.", resolvedPath);
        }

        ExcelDocument document = ExcelDocumentService.LoadDocument(resolvedPath, ReadOnly.IsPresent, AutoSave.IsPresent);
        WriteObject(document);
    }
}
