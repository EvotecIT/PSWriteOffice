using System.IO;
using System.Management.Automation;
using OfficeIMO.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Saves a Word document without disposing it.</summary>
/// <para>Use <c>Close-OfficeWord -Save</c> when you want to save and dispose the document.</para>
/// <example>
///   <summary>Save the open document.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$doc | Save-OfficeWord</code>
///   <para>Persists pending changes and keeps the document open.</para>
/// </example>
[Cmdlet(VerbsData.Save, "OfficeWord")]
[OutputType(typeof(WordDocument))]
public sealed class SaveOfficeWordCommand : PSCmdlet
{
    /// <summary>Document to save.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 0)]
    public WordDocument Document { get; set; } = null!;

    /// <summary>Optional save-as path.</summary>
    [Parameter]
    public string? Path { get; set; }

    /// <summary>Open the document after saving.</summary>
    [Parameter]
    public SwitchParameter Show { get; set; }

    /// <summary>Emit the document object for further processing.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (Document == null)
        {
            return;
        }

        if (string.IsNullOrWhiteSpace(Path) && string.IsNullOrWhiteSpace(Document.FilePath))
        {
            throw new PSInvalidOperationException("No file path provided. Use -Path or open the document from disk.");
        }

        if (!string.IsNullOrWhiteSpace(Path))
        {
            Document.Save(Path!, Show.IsPresent);
        }
        else
        {
            Document.Save(Show.IsPresent);
        }

        if (PassThru.IsPresent)
        {
            WriteObject(Document);
        }
    }
}
