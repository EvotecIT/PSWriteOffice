using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Closes a Word document, optionally saving it.</summary>
/// <para>Provides a cmdlet wrapper for <c>WordDocument.Dispose</c>/<c>Save</c> so scripts need not call .NET methods directly.</para>
/// <example>
///   <summary>Close without saving.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$doc = Get-OfficeWord -Path .\Report.docx; Close-OfficeWord -Document $doc</code>
///   <para>Disposes the loaded document instance without saving changes.</para>
/// </example>
/// <example>
///   <summary>Save to a new path and open the file.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Close-OfficeWord -Document $doc -Save -Path .\Report-final.docx -Show</code>
///   <para>Saves updates to <c>Report-final.docx</c>, opens it, and disposes the document.</para>
/// </example>
[Cmdlet(VerbsCommon.Close, "OfficeWord")]
public sealed class CloseOfficeWordCommand : PSCmdlet
{
    /// <summary>Word document to close.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true)]
    public WordDocument Document { get; set; } = null!;

    /// <summary>Persist changes before closing.</summary>
    [Parameter]
    public SwitchParameter Save { get; set; }

    /// <summary>Optional target path when saving.</summary>
    [Parameter]
    public string? Path { get; set; }

    /// <summary>Open the file after saving.</summary>
    [Parameter]
    public SwitchParameter Show { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (Document == null)
        {
            return;
        }

        if (Save.IsPresent || !string.IsNullOrEmpty(Path))
        {
            WordDocumentService.SaveDocument(Document, Show.IsPresent, Path);
        }
        else
        {
            WordDocumentService.CloseDocument(Document);
        }
    }
}
