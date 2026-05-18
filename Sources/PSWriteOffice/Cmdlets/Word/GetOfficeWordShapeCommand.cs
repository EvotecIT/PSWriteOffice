using System;
using System.Collections.Generic;
using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Gets shapes from a Word document, section, or paragraph.</summary>
[Cmdlet(VerbsCommon.Get, "OfficeWordShape", DefaultParameterSetName = ParameterSetPath)]
[Alias("WordShapes")]
[OutputType(typeof(WordShape))]
public sealed class GetOfficeWordShapeCommand : PSCmdlet
{
    private const string ParameterSetPath = "Path";
    private const string ParameterSetDocument = "Document";
    private const string ParameterSetSection = "Section";
    private const string ParameterSetParagraph = "Paragraph";

    /// <summary>Path to the document.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("FilePath", "Path")]
    public string InputPath { get; set; } = string.Empty;

    /// <summary>Document to inspect.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public WordDocument Document { get; set; } = null!;

    /// <summary>Section to inspect.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetSection)]
    public WordSection Section { get; set; } = null!;

    /// <summary>Paragraph to inspect.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetParagraph)]
    public WordParagraph Paragraph { get; set; } = null!;

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        WordDocument? document = null;
        var dispose = false;

        try
        {
            IEnumerable<WordShape> shapes;
            switch (ParameterSetName)
            {
                case ParameterSetSection:
                    shapes = Section != null
                        ? Section.Shapes
                        : Array.Empty<WordShape>();
                    break;
                case ParameterSetParagraph:
                    shapes = Paragraph?.Shape != null
                        ? new[] { Paragraph.Shape }
                        : Array.Empty<WordShape>();
                    break;
                default:
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

                    shapes = document != null
                        ? document.Shapes
                        : Array.Empty<WordShape>();
                    break;
            }

            WriteObject(shapes, enumerateCollection: true);
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
