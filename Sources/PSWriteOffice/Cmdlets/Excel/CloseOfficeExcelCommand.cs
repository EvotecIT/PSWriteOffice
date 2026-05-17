using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Closes an Excel workbook and optionally saves it.</summary>
/// <para>Convenience wrapper so scripts do not need to call <see cref="ExcelDocument.Save()"/> or <c>Dispose</c> directly.</para>
/// <example>
///   <summary>Save to a new path and open the file.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Close-OfficeExcel -Document $workbook -Save -Path .\report-final.xlsx -Show</code>
///   <para>Saves pending changes to a new file, launches Excel, and releases the workbook.</para>
/// </example>
[Cmdlet(VerbsCommon.Close, "OfficeExcel")]
public sealed class CloseOfficeExcelCommand : PSCmdlet
{
    /// <summary>Workbook to close.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true)]
    public ExcelDocument Document { get; set; } = null!;

    /// <summary>Persist changes before closing.</summary>
    [Parameter]
    public SwitchParameter Save { get; set; }

    /// <summary>Optional output path when saving.</summary>
    [Parameter]
    public string? Path { get; set; }

    /// <summary>Open the workbook in Excel after saving.</summary>
    [Parameter]
    public SwitchParameter Show { get; set; }

    /// <summary>Password used to save the workbook as an encrypted package.</summary>
    [Parameter]
    public string? Password { get; set; }

    /// <summary>Run OfficeIMO worksheet preflight cleanup before saving.</summary>
    [Parameter]
    public SwitchParameter SafePreflight { get; set; }

    /// <summary>Repair common defined-name issues before saving.</summary>
    [Parameter]
    public SwitchParameter SafeRepairDefinedNames { get; set; }

    /// <summary>Validate the saved package with OpenXmlValidator and throw on errors.</summary>
    [Parameter]
    public SwitchParameter ValidateOpenXml { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (Document == null)
        {
            return;
        }

        if (Save.IsPresent || !string.IsNullOrEmpty(Path))
        {
            var resolvedPath = !string.IsNullOrWhiteSpace(Path)
                ? SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path)
                : Document.FilePath;
            var saveOptions = ExcelDocumentService.CreateSaveOptions(
                SafePreflight.IsPresent,
                SafeRepairDefinedNames.IsPresent,
                ValidateOpenXml.IsPresent);
            ExcelDocumentService.SaveDocument(Document, Show.IsPresent, resolvedPath, Password, saveOptions);
        }
        else
        {
            ExcelDocumentService.CloseDocument(Document);
        }
    }
}
