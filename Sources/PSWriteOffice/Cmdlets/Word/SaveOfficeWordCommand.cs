using System.IO;
using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services;

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
    [Alias("FilePath")]
    public string? Path { get; set; }

    /// <summary>Open the document after saving.</summary>
    [Parameter]
    public SwitchParameter Show { get; set; }

    /// <summary>Password used to save the document as an encrypted package.</summary>
    [Parameter]
    public string? Password { get; set; }

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
            var resolvedPath = SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path);
            if (!string.IsNullOrEmpty(Password))
            {
                OfficeEncryptedPackageService.SaveWord(Document, resolvedPath, Password!, false);
            }
            else
            {
                Document.Save(resolvedPath, false);
            }

            if (Show.IsPresent)
            {
                FileOpenService.Open(resolvedPath);
            }
        }
        else
        {
            if (!string.IsNullOrEmpty(Password))
            {
                OfficeEncryptedPackageService.SaveWord(Document, Document.FilePath!, Password!, false);
            }
            else
            {
                Document.Save(false);
            }

            if (Show.IsPresent)
            {
                FileOpenService.Open(Document.FilePath);
            }
        }

        if (PassThru.IsPresent)
        {
            WriteObject(Document);
        }
    }
}
