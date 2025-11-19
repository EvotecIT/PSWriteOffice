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

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (Document == null)
        {
            return;
        }

        if (Save.IsPresent || !string.IsNullOrEmpty(Path))
        {
            ExcelDocumentService.SaveDocument(Document, Show.IsPresent, Path ?? Document.FilePath);
        }
        else
        {
            ExcelDocumentService.CloseDocument(Document);
        }
    }
}
