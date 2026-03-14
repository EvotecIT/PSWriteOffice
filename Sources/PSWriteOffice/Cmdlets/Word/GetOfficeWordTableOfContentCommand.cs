using System;
using System.IO;
using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Gets the table of contents from a Word document.</summary>
/// <example>
///   <summary>Retrieve the table of contents if present.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Get-OfficeWordTableOfContent -Path .\Report.docx</code>
///   <para>Returns the table of contents when one exists.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeWordTableOfContent", DefaultParameterSetName = ParameterSetPath)]
[Alias("WordTableOfContents")]
[OutputType(typeof(WordTableOfContent))]
public sealed class GetOfficeWordTableOfContentCommand : PSCmdlet
{
    private const string ParameterSetPath = "Path";
    private const string ParameterSetDocument = "Document";

    /// <summary>Path to the .docx file.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("FilePath", "Path")]
    public string InputPath { get; set; } = string.Empty;

    /// <summary>Word document to read.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public WordDocument Document { get; set; } = null!;

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        WordDocument? document = null;
        var dispose = false;

        try
        {
            if (ParameterSetName == ParameterSetPath)
            {
                var resolvedPath = SessionState.Path.GetUnresolvedProviderPathFromPSPath(InputPath);
                document = WordDocumentService.LoadDocument(resolvedPath, readOnly: true, autoSave: false);
                dispose = true;
            }
            else
            {
                document = Document;
            }

            if (document == null)
            {
                throw new InvalidOperationException("Word document was not provided.");
            }

            var toc = document.TableOfContent;
            if (toc != null)
            {
                WriteObject(toc);
            }
        }
        finally
        {
            if (dispose)
            {
                document?.Dispose();
            }
        }
    }
}
