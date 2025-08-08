using System;
using System.Management.Automation;
using ClosedXML.Excel;
using PSWriteOffice.Services.Excel;
using ValidateScriptAttribute = PSWriteOffice.Validation.ValidateScriptAttribute;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Saves an Excel workbook to disk.</summary>
/// <para>Writes the provided <see cref="XLWorkbook"/> to the specified file path.</para>
/// <list type="alertSet">
/// <item>
/// <description>Existing files at the destination path are overwritten without confirmation.</description>
/// </item>
/// </list>
/// <example>
/// <summary>Save a workbook to a path</summary>
/// <prefix>PS&gt; </prefix>
/// <code>Save-OfficeExcel -Workbook $wb -FilePath 'report.xlsx'</code>
/// <para>Saves the workbook to the specified location.</para>
/// </example>
/// <example>
/// <summary>Save and open the workbook</summary>
/// <prefix>PS&gt; </prefix>
/// <code>Save-OfficeExcel -Workbook $wb -FilePath 'report.xlsx' -Show</code>
/// <para>After saving, the file is opened in the associated application.</para>
/// </example>
/// <seealso href="https://learn.microsoft.com/en-us/dotnet/api/closedxml.excel.xlworkbook.saveas" />
/// <seealso href="https://github.com/EvotecIT/PSWriteOffice" />
[Cmdlet(VerbsData.Save, "OfficeExcel", SupportsShouldProcess = true)]
public class SaveOfficeExcelCommand : PSCmdlet
{
    /// <summary>Workbook to save.</summary>
    [Parameter(Mandatory = true)]
    public XLWorkbook Workbook { get; set; } = null!;

    /// <summary>Destination path for the workbook file.</summary>
    [Parameter(Mandatory = true)]
    [ValidateNotNullOrEmpty]
    [ValidateScript("{ Test-Path $_ }")]
    public string FilePath { get; set; } = string.Empty;

    /// <summary>Opens the workbook after saving.</summary>
    [Parameter]
    public SwitchParameter Show { get; set; }

    protected override void ProcessRecord()
    {
        try
        {
            if (ShouldProcess(FilePath, "Save workbook"))
            {
                ExcelDocumentService.SaveWorkbook(Workbook, FilePath, Show);
            }
        }
        catch (Exception ex)
        {
            WriteError(new ErrorRecord(ex, "ExcelSaveFailed", ErrorCategory.WriteError, FilePath));
        }
    }
}
