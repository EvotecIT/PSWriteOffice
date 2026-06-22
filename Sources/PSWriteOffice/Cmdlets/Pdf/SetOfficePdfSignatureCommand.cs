using System.IO;
using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Injects externally produced CMS, CAdES, or timestamp signature bytes into a prepared PDF signature placeholder.</summary>
/// <example>
///   <summary>Inject a detached CMS signature into a prepared PDF.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Set-OfficePdfSignature -Path .\Prepared.pdf -SignaturePath .\signature.der -OutputPath .\Signed.pdf
/// Get-OfficePdfSignature -Path .\Signed.pdf</code>
///   <para>Writes a PDF with the reserved /Contents hex slot patched in place.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficePdfSignature")]
[OutputType(typeof(FileInfo), typeof(PdfSignatureValidationReport))]
public sealed class SetOfficePdfSignatureCommand : PSCmdlet
{
    /// <summary>Prepared PDF path.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>DER/CMS/CAdES/TSA response bytes to inject into the reserved /Contents slot.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    public string SignaturePath { get; set; } = string.Empty;

    /// <summary>Output signed PDF path.</summary>
    [Parameter(Mandatory = true, Position = 2)]
    public string OutputPath { get; set; } = string.Empty;

    /// <summary>Return a signature validation report for the written PDF instead of only the output file.</summary>
    [Parameter]
    public SwitchParameter PassThruReport { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var inputPath = PdfCommandUtilities.ResolvePath(this, Path);
        var signaturePath = PdfCommandUtilities.ResolvePath(this, SignaturePath);
        var outputPath = PdfCommandUtilities.ResolvePath(this, OutputPath);
        PdfCommandUtilities.EnsureDirectory(outputPath);

        PdfIncrementalUpdater.ApplyExternalSignature(inputPath, outputPath, File.ReadAllBytes(signaturePath));
        WriteObject(PassThruReport.IsPresent
            ? PdfSignatureValidator.Validate(outputPath)
            : new FileInfo(outputPath));
    }
}
