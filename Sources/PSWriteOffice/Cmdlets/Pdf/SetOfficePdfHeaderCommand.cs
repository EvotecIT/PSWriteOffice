using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Sets running PDF header text.</summary>
/// <example>
///   <summary>Add a running report header.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficePdf -Path .\Examples\Documents\PdfHeader.pdf {
///     Set-OfficePdfHeader -Text 'Service Review' -Align Right -FontSize 9
///     Add-OfficePdfHeading -Text 'Service Review'
///     Add-OfficePdfParagraph -Text 'The header repeats on generated pages.'
/// }</code>
///   <para>Sets header text for the generated PDF.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficePdfHeader", DefaultParameterSetName = ParameterSetContext)]
[Alias("PdfHeader")]
[OutputType(typeof(PdfDocument))]
public sealed class SetOfficePdfHeaderCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";

    /// <summary>PDF document to update outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public PdfDocument Document { get; set; } = null!;

    /// <summary>Header text. Supports {page} and {pages}.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Text { get; set; } = string.Empty;

    /// <summary>Header alignment.</summary>
    [Parameter]
    public PdfAlign Align { get; set; } = PdfAlign.Center;

    /// <summary>Header font size in PDF points.</summary>
    [Parameter]
    public double? FontSize { get; set; }

    /// <summary>Emit the updated document.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var document = PdfCommandUtilities.ResolveDocument(this, Document, ParameterSetName, ParameterSetDocument);
        document.Header(header =>
        {
            ApplyAlignment(header);
            if (FontSize.HasValue)
            {
                header.FontSize(FontSize.Value);
            }
            header.Text(Text);
        });

        if (PassThru.IsPresent)
        {
            WriteObject(document);
        }
    }

    private void ApplyAlignment(PdfHeaderCompose header)
    {
        switch (Align)
        {
            case PdfAlign.Right:
                header.AlignRight();
                break;
            case PdfAlign.Left:
                header.AlignLeft();
                break;
            default:
                header.AlignCenter();
                break;
        }
    }
}
