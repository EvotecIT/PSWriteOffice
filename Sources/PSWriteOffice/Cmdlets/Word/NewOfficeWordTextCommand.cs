/*
.SYNOPSIS
Adds text to a Word document or paragraph.

.DESCRIPTION
Appends text to a Word document or paragraph with optional formatting and alignment.

.EXAMPLE
PS> .\Examples\Word\New-OfficeWordText.ps1
Demonstrates adding formatted text to a document.
*/
using System;
using System.Management.Automation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

[Cmdlet(VerbsCommon.New, "OfficeWordText", DefaultParameterSetName = "Document")]
public class NewOfficeWordTextCommand : PSCmdlet
{
    [Parameter(ParameterSetName = "Document", Mandatory = true)]
    public WordDocument Document { get; set; } = null!;

    [Parameter(ParameterSetName = "Paragraph", Mandatory = true)]
    public WordParagraph Paragraph { get; set; } = null!;

    [Parameter(ParameterSetName = "Document")]
    [Parameter(ParameterSetName = "Paragraph")]
    public string[] Text { get; set; } = Array.Empty<string>();

    [Parameter(ParameterSetName = "Document")]
    [Parameter(ParameterSetName = "Paragraph")]
    public bool?[]? Bold { get; set; }

    [Parameter(ParameterSetName = "Document")]
    [Parameter(ParameterSetName = "Paragraph")]
    public bool?[]? Italic { get; set; }

    [Parameter(ParameterSetName = "Document")]
    [Parameter(ParameterSetName = "Paragraph")]
    public UnderlineValues?[]? Underline { get; set; }

    [Parameter(ParameterSetName = "Document")]
    [Parameter(ParameterSetName = "Paragraph")]
    public string[]? Color { get; set; }

    [Parameter(ParameterSetName = "Document")]
    [Parameter(ParameterSetName = "Paragraph")]
    public JustificationValues? Alignment { get; set; }

    [Parameter(ParameterSetName = "Document")]
    [Parameter(ParameterSetName = "Paragraph")]
    public WordParagraphStyles? Style { get; set; }

    [Parameter(ParameterSetName = "Document")]
    [Parameter(ParameterSetName = "Paragraph")]
    public SwitchParameter ReturnObject { get; set; }

    protected override void ProcessRecord()
    {
        var paragraph = WordDocumentService.AddText(
            ParameterSetName == "Document" ? Document : null,
            ParameterSetName == "Paragraph" ? Paragraph : null,
            Text,
            Bold,
            Italic,
            Underline,
            Color,
            Alignment,
            Style);

        if (ReturnObject.IsPresent)
        {
            WriteObject(paragraph);
        }
    }
}
