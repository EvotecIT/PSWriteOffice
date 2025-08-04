using System;
using System.IO;
using System.Management.Automation;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

[Cmdlet(VerbsData.ConvertFrom, "HTMLtoWord", DefaultParameterSetName = "HTML")]
public class ConvertFromHTMLtoWordCommand : PSCmdlet
{
    [Parameter(Mandatory = true, ParameterSetName = "HTMLFile")]
    [Parameter(Mandatory = true, ParameterSetName = "HTML")]
    public string OutputFile { get; set; } = string.Empty;

    [Parameter(Mandatory = true, ParameterSetName = "HTMLFile")]
    [Alias("InputFile")]
    public string FileHTML { get; set; } = string.Empty;

    [Parameter(Mandatory = true, ParameterSetName = "HTML")]
    [Alias("HTML")]
    public string SourceHTML { get; set; } = string.Empty;

    [Parameter(ParameterSetName = "HTMLFile")]
    [Parameter(ParameterSetName = "HTML")]
    public SwitchParameter Show { get; set; }

    [Parameter(ParameterSetName = "HTMLFile")]
    [Parameter(ParameterSetName = "HTML")]
    public HtmlImportMode Mode { get; set; } = HtmlImportMode.Parse;

    protected override void ProcessRecord()
    {
        string html;
        if (this.ParameterSetName == "HTMLFile")
        {
            html = File.ReadAllText(FileHTML);
        }
        else
        {
            html = SourceHTML;
        }

        try
        {
            var document = WordDocumentService.CreateDocument(OutputFile, false);
            WordDocumentService.AddHtml(document, html, Mode);
            WordDocumentService.SaveDocument(document, Show.IsPresent, null);
        }
        catch (Exception ex)
        {
            WriteError(new ErrorRecord(ex, "HtmlConversionFailed", ErrorCategory.InvalidOperation, OutputFile));
        }
    }
}
