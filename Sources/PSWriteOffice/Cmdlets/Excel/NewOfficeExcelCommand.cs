using System.Management.Automation;
using ClosedXML.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Creates a new in-memory Excel workbook.</summary>
/// <para>Use this cmdlet to start building an Excel document before adding worksheets and data.</para>
/// <list type="alertSet">
/// <item>
/// <description>The workbook exists only in memory until saved with Save-OfficeExcel.</description>
/// </item>
/// </list>
/// <example>
/// <summary>Create a blank workbook</summary>
/// <prefix>PS&gt; </prefix>
/// <code>New-OfficeExcel</code>
/// <para>Returns an empty <c>XLWorkbook</c> instance for further manipulation.</para>
/// </example>
/// <example>
/// <summary>Create and save a workbook</summary>
/// <prefix>PS&gt; </prefix>
/// <code>$wb = New-OfficeExcel; Save-OfficeExcel -Workbook $wb -FilePath 'report.xlsx'</code>
/// <para>The workbook is created and immediately persisted to disk.</para>
/// </example>
/// <seealso href="https://learn.microsoft.com/en-us/dotnet/api/closedxml.excel.xlworkbook" />
/// <seealso href="https://github.com/EvotecIT/PSWriteOffice" />
[Cmdlet(VerbsCommon.New, "OfficeExcel")]
public class NewOfficeExcelCommand : PSCmdlet
{
    protected override void ProcessRecord()
    {
        var workbook = ExcelDocumentService.CreateWorkbook();
        WriteObject(workbook);
    }
}
