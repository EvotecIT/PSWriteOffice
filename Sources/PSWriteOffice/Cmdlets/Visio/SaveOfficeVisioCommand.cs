using System.IO;
using System.Management.Automation;
using OfficeIMO.Visio;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Visio;

namespace PSWriteOffice.Cmdlets.Visio;

/// <summary>Saves an OfficeIMO.Visio document.</summary>
[Cmdlet(VerbsData.Save, "OfficeVisio")]
[Alias("VisioSave")]
[OutputType(typeof(VisioDocument), typeof(FileInfo))]
public sealed class SaveOfficeVisioCommand : PSCmdlet
{
    /// <summary>Visio document to save.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 0)]
    public VisioDocument Document { get; set; } = null!;

    /// <summary>Optional save-as path.</summary>
    [Parameter(Position = 1)]
    [Alias("FilePath")]
    public string? Path { get; set; }

    /// <summary>Open the document after saving.</summary>
    [Parameter]
    public SwitchParameter Show { get; set; }

    /// <summary>Emit the document object instead of the saved file.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (string.IsNullOrWhiteSpace(Path))
        {
            Document.Save();
            WriteObject(Document);
            return;
        }

        var fullPath = VisioCommandUtilities.ResolvePath(this, Path!);
        VisioCommandUtilities.EnsureDirectory(fullPath);
        Document.Save(fullPath);

        if (Show.IsPresent)
        {
            FileOpenService.Open(fullPath);
        }

        WriteObject(PassThru.IsPresent ? Document : new FileInfo(fullPath));
    }
}
