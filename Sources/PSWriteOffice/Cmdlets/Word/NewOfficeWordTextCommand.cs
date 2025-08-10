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
        var textLength = Text.Length;
        if (Bold != null && Bold.Length != textLength)
        {
            throw new ArgumentException("Bold length must match Text length.", nameof(Bold));
        }
        if (Italic != null && Italic.Length != textLength)
        {
            throw new ArgumentException("Italic length must match Text length.", nameof(Italic));
        }
        if (Underline != null && Underline.Length != textLength)
        {
            throw new ArgumentException("Underline length must match Text length.", nameof(Underline));
        }
        if (Color != null && Color.Length != textLength)
        {
            throw new ArgumentException("Color length must match Text length.", nameof(Color));
        }

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
