using System;
using System.IO;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Creates a new Excel workbook using the DSL.</summary>
/// <para>Runs the provided script block inside an <c>ExcelSheet</c>/<c>ExcelCell</c> DSL context and saves the file.</para>
/// <example>
///   <summary>Create a workbook with a sheet and a few cells.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficeExcel -Path .\report.xlsx { ExcelSheet 'Data' { ExcelCell -Address 'A1' -Value 'Region' } }</code>
///   <para>Creates <c>report.xlsx</c> and writes “Region” into cell A1 on the Data worksheet.</para>
/// </example>
[Cmdlet(VerbsCommon.New, "OfficeExcel")]
public sealed class NewOfficeExcelCommand : PSCmdlet
{
    /// <summary>Destination path for the workbook.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    [Alias("Path")]
    public string FilePath { get; set; } = string.Empty;

    /// <summary>DSL scriptblock describing workbook content.</summary>
    [Parameter(Position = 1)]
    public ScriptBlock? Content { get; set; }

    /// <summary>Opt into OfficeIMO automatic saves during operations.</summary>
    [Parameter]
    public SwitchParameter AutoSave { get; set; }

    /// <summary>Skip saving the workbook after running the DSL.</summary>
    [Parameter]
    public SwitchParameter NoSave { get; set; }

    /// <summary>Open the workbook in Excel after saving.</summary>
    [Parameter]
    public SwitchParameter Open { get; set; }

    /// <summary>Emit a <see cref="FileInfo"/> for convenience.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var resolvedPath = SessionState.Path.GetUnresolvedProviderPathFromPSPath(FilePath);
        var directory = Path.GetDirectoryName(resolvedPath);
        if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
        {
            Directory.CreateDirectory(directory);
        }

        var document = ExcelDocumentService.CreateDocument(resolvedPath, AutoSave.IsPresent);
        try
        {
            using (ExcelDslContext.Enter(document))
            {
                Content?.InvokeReturnAsIs();
            }

            if (!NoSave.IsPresent)
            {
                if (document.Sheets.Count == 0)
                {
                    document.AddWorkSheet(string.Empty, SheetNameValidationMode.Sanitize);
                }
                ExcelDocumentService.SaveDocument(document, Open.IsPresent, resolvedPath);
            }
            else
            {
                document.Dispose();
            }
        }
        catch
        {
            document.Dispose();
            throw;
        }

        if (PassThru.IsPresent)
        {
            WriteObject(new FileInfo(resolvedPath));
        }
    }
}
