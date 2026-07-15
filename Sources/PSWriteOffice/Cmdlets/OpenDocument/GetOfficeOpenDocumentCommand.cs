using System.Management.Automation;
using OfficeIMO.OpenDocument;

namespace PSWriteOffice.Cmdlets.OpenDocument;

/// <summary>Loads a native ODT, ODS, or ODP document.</summary>
[Cmdlet(VerbsCommon.Get, "OfficeOpenDocument")]
[OutputType(typeof(OdfDocument), typeof(OdtDocument), typeof(OdsDocument), typeof(OdpPresentation))]
public sealed class GetOfficeOpenDocumentCommand : PSCmdlet
{
    /// <summary>Path to an ODT, ODS, or ODP file.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true)]
    public string Path { get; set; } = string.Empty;

    /// <summary>Optional bounded package and XML settings.</summary>
    [Parameter]
    public OdfLoadOptions? Options { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() => WriteObject(OdfDocument.Load(
        SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path), Options));
}
