using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Gets date picker content controls from a Word document.</summary>
/// <example>
///   <summary>List date picker controls.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Get-OfficeWordDatePicker -Path .\Report.docx</code>
///   <para>Returns all date picker controls in the document.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeWordDatePicker", DefaultParameterSetName = ParameterSetPath)]
[Alias("WordDatePickers")]
[OutputType(typeof(WordDatePicker))]
public sealed class GetOfficeWordDatePickerCommand : PSCmdlet
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

    /// <summary>Filter by date picker alias (wildcards supported).</summary>
    [Parameter]
    [SupportsWildcards]
    public string[]? Alias { get; set; }

    /// <summary>Filter by date picker tag (wildcards supported).</summary>
    [Parameter]
    [SupportsWildcards]
    public string[]? Tag { get; set; }

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

            var aliasPatterns = WordFilterHelpers.BuildPatterns(Alias);
            var tagPatterns = WordFilterHelpers.BuildPatterns(Tag);

            IEnumerable<WordDatePicker> results = document.DatePickers;
            results = results.Where(control =>
                WordFilterHelpers.Matches(control.Alias, aliasPatterns) &&
                WordFilterHelpers.Matches(control.Tag, tagPatterns));

            WriteObject(results, enumerateCollection: true);
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
