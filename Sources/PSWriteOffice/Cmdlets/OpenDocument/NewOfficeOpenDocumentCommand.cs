using System.IO;
using System.Management.Automation;
using OfficeIMO.OpenDocument;

namespace PSWriteOffice.Cmdlets.OpenDocument;

/// <summary>Creates a native ODT, ODS, or ODP document.</summary>
[Cmdlet(VerbsCommon.New, "OfficeOpenDocument", SupportsShouldProcess = true)]
[OutputType(typeof(OdfDocument), typeof(FileInfo))]
public sealed class NewOfficeOpenDocumentCommand : PSCmdlet {
    /// <summary>OpenDocument text, spreadsheet, or presentation kind.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public OdfDocumentKind Kind { get; set; }

    /// <summary>Optional initial destination path.</summary>
    [Parameter(Position = 1)]
    public string? Path { get; set; }

    /// <summary>Emit the saved file when a destination path is supplied.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        OdfDocument document = Kind switch {
            OdfDocumentKind.Text => OdtDocument.Create(),
            OdfDocumentKind.Spreadsheet => OdsDocument.Create(),
            OdfDocumentKind.Presentation => OdpPresentation.Create(),
            _ => throw new PSArgumentOutOfRangeException(nameof(Kind), Kind, "Use Text, Spreadsheet, or Presentation.")
        };
        if (!string.IsNullOrWhiteSpace(Path)) {
            var path = SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path!);
            OpenDocumentCommandUtilities.ValidateOpenDocumentExtension(path, Kind, nameof(Path));
            if (!ShouldProcess(path, "Create OpenDocument package")) {
                return;
            }
            Directory.CreateDirectory(System.IO.Path.GetDirectoryName(path) ?? SessionState.Path.CurrentFileSystemLocation.Path);
            document.Save(path);
            if (PassThru.IsPresent) {
                WriteObject(new FileInfo(path));
            }
            return;
        }
        WriteObject(document);
    }
}