using System.Management.Automation;
using OfficeIMO.Reader;
using PSWriteOffice.Services.Reader;

namespace PSWriteOffice.Cmdlets.Reader;

/// <summary>Lists OfficeIMO.Reader capabilities registered in the current PSWriteOffice process.</summary>
/// <example>
///   <summary>Show registered Reader adapters.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$capabilities = Get-OfficeDocumentCapability
/// $capabilities | Sort-Object Id | Select-Object Id, Extensions</code>
///   <para>Lists built-in and modular Reader handlers, including adapters such as PDF, RTF, HTML, CSV, JSON, XML, YAML, ZIP, EPUB, and Visio when available.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeDocumentCapability")]
[Alias("Get-OfficeReaderCapability")]
[OutputType(typeof(ReaderHandlerCapability), typeof(ReaderCapabilityManifest))]
public sealed class GetOfficeDocumentCapabilityCommand : PSCmdlet
{
    /// <summary>Return the capability manifest envelope instead of individual handlers.</summary>
    [Parameter]
    public SwitchParameter Manifest { get; set; }

    /// <summary>Exclude built-in Reader capabilities.</summary>
    [Parameter]
    public SwitchParameter ExcludeBuiltIn { get; set; }

    /// <summary>Exclude custom or modular Reader capabilities.</summary>
    [Parameter]
    public SwitchParameter ExcludeCustom { get; set; }

    /// <inheritdoc />
    protected override void BeginProcessing()
    {
        ReaderCommandUtilities.RegisterReaderAdapters();
    }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var includeBuiltIn = !ExcludeBuiltIn.IsPresent;
        var includeCustom = !ExcludeCustom.IsPresent;

        if (Manifest.IsPresent)
        {
            WriteObject(DocumentReader.GetCapabilityManifest(includeBuiltIn, includeCustom));
            return;
        }

        WriteObject(DocumentReader.GetCapabilities(includeBuiltIn, includeCustom), enumerateCollection: true);
    }
}
