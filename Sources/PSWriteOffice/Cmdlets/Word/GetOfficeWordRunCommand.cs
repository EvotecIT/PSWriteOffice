using System;
using System.Collections.Generic;
using System.IO;
using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Gets runs from Word paragraphs.</summary>
/// <example>
///   <summary>Enumerate runs for all paragraphs.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Get-OfficeWordParagraph -Path .\Report.docx | Get-OfficeWordRun</code>
///   <para>Returns each run as a <see cref="WordParagraph"/> instance.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeWordRun", DefaultParameterSetName = ParameterSetParagraph)]
[OutputType(typeof(WordParagraph))]
public sealed class GetOfficeWordRunCommand : PSCmdlet
{
    private const string ParameterSetParagraph = "Paragraph";
    private const string ParameterSetSection = "Section";
    private const string ParameterSetDocument = "Document";
    private const string ParameterSetPath = "Path";

    /// <summary>Paragraph to enumerate.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetParagraph)]
    public WordParagraph Paragraph { get; set; } = null!;

    /// <summary>Section to enumerate.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetSection)]
    public WordSection Section { get; set; } = null!;

    /// <summary>Document to enumerate.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public WordDocument Document { get; set; } = null!;

    /// <summary>Path to the document.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("FilePath", "Path")]
    public string InputPath { get; set; } = string.Empty;

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        WordDocument? document = null;
        var dispose = false;

        try
        {
            IEnumerable<WordParagraph> paragraphs;

            switch (ParameterSetName)
            {
                case ParameterSetParagraph:
                    paragraphs = Paragraph != null ? new[] { Paragraph } : Array.Empty<WordParagraph>();
                    break;
                case ParameterSetSection:
                    paragraphs = Section != null
                        ? Section.Paragraphs
                        : Array.Empty<WordParagraph>();
                    break;
                case ParameterSetPath:
                    var resolvedPath = SessionState.Path.GetUnresolvedProviderPathFromPSPath(InputPath);
                    document = WordDocumentService.LoadDocument(resolvedPath, readOnly: true, autoSave: false);
                    dispose = true;
                    paragraphs = document.Paragraphs;
                    break;
                default:
                    document = Document;
                    if (document == null)
                    {
                        throw new InvalidOperationException("Word document was not provided.");
                    }
                    paragraphs = document.Paragraphs;
                    break;
            }

            foreach (var paragraph in paragraphs)
            {
                foreach (var run in paragraph.GetRuns())
                {
                    WriteObject(run);
                }
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
