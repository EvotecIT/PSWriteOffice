using System;
using System.IO;
using System.Management.Automation;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Creates a Word document using the DSL.</summary>
/// <para>Handles file creation, scriptblock execution, optional autosave, and emits the document path when <c>-PassThru</c> is used.</para>
/// <example>
///   <summary>Create a document inline.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficeWord -Path .\Report.docx { WordSection { WordParagraph 'Hello DSL' } } -Open</code>
///   <para>Builds a document, adds one paragraph, saves it to disk, and opens it.</para>
/// </example>
[Cmdlet(VerbsCommon.New, "OfficeWord")]
public sealed class NewOfficeWordCommand : PSCmdlet
{
    /// <summary>Destination path for the document.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    [Alias("FilePath", "Path")]
    public string OutputPath { get; set; } = string.Empty;

    /// <summary>DSL scriptblock describing document content.</summary>
    [Parameter(Position = 1)]
    public ScriptBlock? Content { get; set; }

    /// <summary>Emit a <see cref="FileInfo"/> for chaining.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <summary>Open the document after saving.</summary>
    [Parameter]
    public SwitchParameter Open { get; set; }

    /// <summary>Skip saving after executing the DSL.</summary>
    [Parameter]
    public SwitchParameter NoSave { get; set; }

    /// <summary>Enable OfficeIMO AutoSave mode.</summary>
    [Parameter]
    public SwitchParameter AutoSave { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var fullPath = GetResolvedPath();
        var directory = Path.GetDirectoryName(fullPath);
        if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
        {
            Directory.CreateDirectory(directory);
        }

        var document = WordDocumentService.CreateDocument(fullPath, AutoSave.IsPresent);

        if (Content == null)
        {
            WriteObject(document);
            return;
        }

        using (WordDslContext.Enter(document))
        {
            Content.InvokeReturnAsIs();
        }

        if (NoSave.IsPresent)
        {
            document.Dispose();
        }
        else
        {
            WordDocumentService.SaveDocument(document, Open.IsPresent, fullPath);
        }

        if (PassThru.IsPresent)
        {
            WriteObject(new FileInfo(fullPath));
        }
    }

    private string GetResolvedPath()
    {
        var providerPath = SessionState.Path.GetUnresolvedProviderPathFromPSPath(OutputPath);
        return Path.IsPathRooted(providerPath)
            ? providerPath
            : Path.Combine(SessionState.Path.CurrentFileSystemLocation.Path, providerPath);
    }
}
