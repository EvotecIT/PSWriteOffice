using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Sets running PDF footer text.</summary>
/// <example>
///   <summary>Add page numbers to a generated PDF.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficePdf -Path .\Examples\Documents\PdfFooter.pdf {
///     Set-OfficePdfFooter -Text 'Page {page} of {pages}' -Align Center -FontSize 8
///     Add-OfficePdfHeading -Text 'Report with footer'
///     Add-OfficePdfPageBreak
///     Add-OfficePdfParagraph -Text 'The footer includes generated page numbers.'
/// }</code>
///   <para>Uses page placeholders in a running footer.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficePdfFooter", DefaultParameterSetName = ParameterSetContext)]
[Alias("PdfFooter")]
[OutputType(typeof(PdfDocument))]
public sealed class SetOfficePdfFooterCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";

    /// <summary>PDF document to update outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public PdfDocument Document { get; set; } = null!;

    /// <summary>Footer text. Supports {page} and {pages}.</summary>
    [Parameter(Position = 0)]
    public string Text { get; set; } = "Page {page}/{pages}";

    /// <summary>Footer alignment.</summary>
    [Parameter]
    public PdfAlign Align { get; set; } = PdfAlign.Center;

    /// <summary>Footer font size in PDF points.</summary>
    [Parameter]
    public double? FontSize { get; set; }

    /// <summary>Emit the updated document.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var document = PdfCommandUtilities.ResolveDocument(this, Document, ParameterSetName, ParameterSetDocument);
        document.Footer(footer =>
        {
            ApplyAlignment(footer);
            if (FontSize.HasValue)
            {
                footer.FontSize(FontSize.Value);
            }
            footer.Text(Text);
        });

        if (PassThru.IsPresent)
        {
            WriteObject(document);
        }
    }

    private void ApplyAlignment(PdfFooterCompose footer)
    {
        switch (Align)
        {
            case PdfAlign.Right:
                footer.AlignRight();
                break;
            case PdfAlign.Left:
                footer.AlignLeft();
                break;
            default:
                footer.AlignCenter();
                break;
        }
    }
}
