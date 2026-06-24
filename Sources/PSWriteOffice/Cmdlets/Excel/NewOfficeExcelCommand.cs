using System;
using System.IO;
using System.Management.Automation;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using PSWriteOffice.Services.Excel;
using PSWriteOffice.Services.Pdf;

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

    /// <summary>Optional workbook template package copied before running the DSL.</summary>
    [Parameter]
    [Alias("Template")]
    public string? TemplatePath { get; set; }

    /// <summary>Opt into OfficeIMO automatic saves during operations.</summary>
    [Parameter]
    public SwitchParameter AutoSave { get; set; }

    /// <summary>Skip saving the workbook after running the DSL.</summary>
    [Parameter]
    public SwitchParameter NoSave { get; set; }

    /// <summary>Open the workbook in Excel after saving.</summary>
    [Parameter]
    public SwitchParameter Open { get; set; }

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

    /// <summary>Disable OfficeIMO fast package writers for this save.</summary>
    [Parameter]
    public SwitchParameter DisableFastPackageWriter { get; set; }

    /// <summary>Evaluate supported formulas and write cached values before saving.</summary>
    [Parameter]
    public SwitchParameter EvaluateFormulas { get; set; }

    /// <summary>Remove cached formula results before saving.</summary>
    [Parameter]
    public SwitchParameter ClearCachedFormulaResults { get; set; }

    /// <summary>Mark formula cells dirty before saving.</summary>
    [Parameter]
    public SwitchParameter MarkFormulasDirty { get; set; }

    /// <summary>Request a full workbook recalculation when opened in Excel-compatible applications.</summary>
    [Parameter]
    public SwitchParameter ForceFullCalculationOnOpen { get; set; }

    /// <summary>Workbook date system for Excel date serials.</summary>
    [Parameter]
    [ValidateSet("1900", "1904", "NineteenHundred", "NineteenFour")]
    public string? DateSystem { get; set; }

    /// <summary>Optional PDF path to create from the same workbook before closing it.</summary>
    [Parameter]
    public string? PdfPath { get; set; }

    /// <summary>Emit a <see cref="FileInfo"/> for convenience.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <summary>Workbook document title metadata.</summary>
    [Parameter]
    public string? DocumentTitle { get; set; }

    /// <summary>Workbook author metadata.</summary>
    [Parameter]
    public string? Author { get; set; }

    /// <summary>Workbook subject metadata.</summary>
    [Parameter]
    public string? Subject { get; set; }

    /// <summary>Workbook keyword metadata.</summary>
    [Parameter]
    public string? Keywords { get; set; }

    /// <summary>Workbook description metadata.</summary>
    [Parameter]
    public string? Description { get; set; }

    /// <summary>Workbook category metadata.</summary>
    [Parameter]
    public string? Category { get; set; }

    /// <summary>Workbook company metadata.</summary>
    [Parameter]
    public string? Company { get; set; }

    /// <summary>Workbook manager metadata.</summary>
    [Parameter]
    public string? Manager { get; set; }

    /// <summary>Workbook application-name metadata.</summary>
    [Parameter]
    public string? ApplicationName { get; set; }

    /// <summary>Workbook last-modified-by metadata.</summary>
    [Parameter]
    public string? LastModifiedBy { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var resolvedPath = SessionState.Path.GetUnresolvedProviderPathFromPSPath(FilePath);
        var directory = Path.GetDirectoryName(resolvedPath);
        if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
        {
            Directory.CreateDirectory(directory);
        }

        var document = string.IsNullOrWhiteSpace(TemplatePath)
            ? ExcelDocumentService.CreateDocument(resolvedPath, AutoSave.IsPresent)
            : ExcelDocumentService.CreateDocumentFromTemplate(
                SessionState.Path.GetUnresolvedProviderPathFromPSPath(TemplatePath!),
                resolvedPath,
                AutoSave.IsPresent);
        try
        {
            ExcelDateSystemService.ApplyIfSpecified(document, DateSystem, nameof(DateSystem));
            ExcelDocumentPropertyService.ApplyCommonProperties(
                document,
                DocumentTitle,
                Author,
                Subject,
                Keywords,
                Description,
                Category,
                Company,
                Manager,
                ApplicationName,
                LastModifiedBy);

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
                var saveOptions = ExcelDocumentService.CreateSaveOptions(
                    SafePreflight.IsPresent,
                    SafeRepairDefinedNames.IsPresent,
                    ValidateOpenXml.IsPresent,
                    DisableFastPackageWriter.IsPresent,
                    EvaluateFormulas.IsPresent,
                    ClearCachedFormulaResults.IsPresent,
                    MarkFormulasDirty.IsPresent,
                    ForceFullCalculationOnOpen.IsPresent);
                SavePdfIfRequested(document);
                ExcelDocumentService.SaveDocument(document, Open.IsPresent, resolvedPath, Password, saveOptions);
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

    private void SavePdfIfRequested(ExcelDocument document)
    {
        if (string.IsNullOrWhiteSpace(PdfPath))
        {
            return;
        }

        var pdfPath = PdfCommandUtilities.ResolvePath(this, PdfPath!);
        PdfCommandUtilities.EnsureDirectory(pdfPath);
        document.SaveAsPdf(pdfPath);
    }
}
