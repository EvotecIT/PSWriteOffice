using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Gets simple AcroForm fields from a PDF.</summary>
[Cmdlet(VerbsCommon.Get, "OfficePdfFormField")]
[OutputType(typeof(PdfFormField))]
public sealed class GetOfficePdfFormFieldCommand : PSCmdlet
{
    /// <summary>PDF file path.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Optional field name filter.</summary>
    [Parameter]
    public string? Name { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var fields = PdfInspector.Inspect(PdfCommandUtilities.ResolvePath(this, Path)).FormFields;
        foreach (var field in fields)
        {
            if (string.IsNullOrWhiteSpace(Name) || string.Equals(field.Name, Name, System.StringComparison.Ordinal))
            {
                WriteObject(field);
            }
        }
    }
}
