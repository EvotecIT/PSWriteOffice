using System.Management.Automation;
using OfficeIMO.Visio;
using PSWriteOffice.Services.Visio;

namespace PSWriteOffice.Cmdlets.Visio;

/// <summary>Creates a deterministic inspection snapshot for a Visio document.</summary>
[Cmdlet(VerbsCommon.Get, "OfficeVisioInfo", DefaultParameterSetName = PathParameterSet)]
[Alias("VisioInfo")]
[OutputType(typeof(VisioInspectionSnapshot), typeof(string))]
public sealed class GetOfficeVisioInfoCommand : PSCmdlet
{
    private const string PathParameterSet = "Path";
    private const string DocumentParameterSet = "Document";

    /// <summary>Visio .vsdx path.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = PathParameterSet)]
    [Alias("FilePath")]
    public string? Path { get; set; }

    /// <summary>Visio document object.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = DocumentParameterSet)]
    public VisioDocument? Document { get; set; }

    /// <summary>Emit the stable line-oriented inspection text.</summary>
    [Parameter]
    public SwitchParameter AsText { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var document = VisioCommandUtilities.ResolveDocument(this, Document, Path);
        var snapshot = document.CreateInspectionSnapshot();
        WriteObject(AsText.IsPresent ? snapshot.ToText() : snapshot);
    }
}
