using System.Management.Automation;
using OfficeIMO.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Saves an Excel workbook without disposing it.</summary>
/// <example>
///   <summary>Save a workbook in-place.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$workbook | Save-OfficeExcel</code>
///   <para>Writes pending changes to disk and keeps the workbook open.</para>
/// </example>
[Cmdlet(VerbsData.Save, "OfficeExcel")]
[OutputType(typeof(ExcelDocument))]
public sealed class SaveOfficeExcelCommand : PSCmdlet
{
    /// <summary>Workbook to save.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 0)]
    public ExcelDocument Document { get; set; } = null!;

    /// <summary>Optional save-as path.</summary>
    [Parameter]
    public string? Path { get; set; }

    /// <summary>Open the workbook after saving.</summary>
    [Parameter]
    public SwitchParameter Show { get; set; }

    /// <summary>Emit the workbook for further processing.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (Document == null)
        {
            return;
        }

        if (string.IsNullOrWhiteSpace(Path) && string.IsNullOrWhiteSpace(Document.FilePath))
        {
            throw new PSInvalidOperationException("No file path provided. Use -Path or open the workbook from disk.");
        }

        if (!string.IsNullOrWhiteSpace(Path))
        {
            Document.Save(Path!, Show.IsPresent);
        }
        else
        {
            Document.Save(Show.IsPresent);
        }

        if (PassThru.IsPresent)
        {
            WriteObject(Document);
        }
    }
}
