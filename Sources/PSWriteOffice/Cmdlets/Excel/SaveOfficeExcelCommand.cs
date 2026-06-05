using System.Management.Automation;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Excel;
using PSWriteOffice.Services.Pdf;

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

    /// <summary>Optional PDF path to create from the same workbook.</summary>
    [Parameter]
    public string? PdfPath { get; set; }

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

        var saveOptions = ExcelDocumentService.CreateSaveOptions(
            SafePreflight.IsPresent,
            SafeRepairDefinedNames.IsPresent,
            ValidateOpenXml.IsPresent);

        if (!string.IsNullOrWhiteSpace(Path))
        {
            var resolvedPath = SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path);
            if (!string.IsNullOrEmpty(Password))
            {
                OfficeEncryptedPackageService.SaveExcel(Document, resolvedPath, Password!, false, saveOptions);
            }
            else
            {
                Document.Save(resolvedPath, false, saveOptions);
            }

            if (Show.IsPresent)
            {
                FileOpenService.Open(resolvedPath);
            }
        }
        else
        {
            if (!string.IsNullOrEmpty(Password))
            {
                OfficeEncryptedPackageService.SaveExcel(Document, Document.FilePath!, Password!, false, saveOptions);
            }
            else
            {
                if (saveOptions == null)
                {
                    Document.Save(false);
                }
                else
                {
                    Document.Save(Document.FilePath!, false, saveOptions);
                }
            }

            if (Show.IsPresent)
            {
                FileOpenService.Open(Document.FilePath);
            }
        }

        SavePdfIfRequested();

        if (PassThru.IsPresent)
        {
            WriteObject(Document);
        }
    }

    private void SavePdfIfRequested()
    {
        if (string.IsNullOrWhiteSpace(PdfPath))
        {
            return;
        }

        Document.SaveAsPdf(PdfCommandUtilities.ResolvePath(this, PdfPath!));
    }
}
