using System.IO;
using System.Management.Automation;
using OfficeIMO.Email;

namespace PSWriteOffice.Cmdlets.Email;

/// <summary>Saves an email document as EML, MSG, or TNEF with fidelity diagnostics.</summary>
[Cmdlet(VerbsData.Save, "OfficeEmail", SupportsShouldProcess = true)]
[OutputType(typeof(EmailWriteResult))]
public sealed class SaveOfficeEmailCommand : PSCmdlet
{
    /// <summary>Email document to save.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true)]
    public EmailDocument Document { get; set; } = null!;

    /// <summary>Destination path.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Path { get; set; } = string.Empty;

    /// <summary>Optional explicit output format. By default it is inferred from the filename.</summary>
    [Parameter]
    public EmailFileFormat? Format { get; set; }

    /// <summary>Optional preservation, projection, encoding, and output limits.</summary>
    [Parameter]
    public EmailWriterOptions? Options { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var output = SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path);
        if (!ShouldProcess(output, "Save email artifact")) return;
        Directory.CreateDirectory(System.IO.Path.GetDirectoryName(output) ?? SessionState.Path.CurrentFileSystemLocation.Path);
        WriteObject(Format.HasValue
            ? Document.Save(output, Format.Value, Options)
            : Document.Save(output, Options));
    }
}
