using System.Collections;
using System.IO;
using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Fills and optionally flattens simple AcroForm fields in an existing PDF.</summary>
/// <example>
///   <summary>Fill and flatten a PDF form.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$fields = @{
///     Requester = 'Ada Lovelace'
///     Priority = 'High'
///     Approved = $true
/// }
/// Set-OfficePdfForm -Path .\Examples\Documents\Request.pdf -OutputPath .\Examples\Documents\Request-FilledFlat.pdf -Field $fields -Flatten</code>
///   <para>Fills simple AcroForm fields and writes a flattened PDF.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficePdfForm")]
[OutputType(typeof(FileInfo))]
public sealed class SetOfficePdfFormCommand : PSCmdlet
{
    /// <summary>Input PDF path.</summary>
    [Parameter(Mandatory = true)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Output PDF path.</summary>
    [Parameter(Mandatory = true)]
    public string OutputPath { get; set; } = string.Empty;

    /// <summary>Field values keyed by form field name.</summary>
    [Parameter]
    public Hashtable? Field { get; set; }

    /// <summary>Flatten simple form fields after filling, or flatten without filling when -Field is omitted.</summary>
    [Parameter]
    public SwitchParameter Flatten { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var document = PdfDocument.Open(PdfCommandUtilities.ResolvePath(this, Path));
        PdfDocument result;
        if (Field == null || Field.Count == 0)
        {
            if (!Flatten.IsPresent)
            {
                throw new PSArgumentException("Provide -Field values or use -Flatten.", nameof(Field));
            }

            result = document.Forms.Flatten();
        }
        else
        {
            var values = PdfCommandUtilities.ConvertFieldValues(Field);
            result = Flatten.IsPresent
                ? document.Forms.FillAndFlatten(values)
                : document.Forms.Fill(values);
        }

        var outputPath = PdfCommandUtilities.ResolvePath(this, OutputPath);
        PdfCommandUtilities.EnsureDirectory(outputPath);
        result.Save(outputPath);
        WriteObject(new FileInfo(outputPath));
    }
}
