using System.Collections;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Adds a table to a PDF document.</summary>
[Cmdlet(VerbsCommon.Add, "OfficePdfTable", DefaultParameterSetName = ParameterSetContext)]
[Alias("PdfTable")]
[OutputType(typeof(PdfDocument))]
public sealed class AddOfficePdfTableCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";

    /// <summary>PDF document to update outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public PdfDocument Document { get; set; } = null!;

    /// <summary>Objects or row arrays to render as a table.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public object[] InputObject { get; set; } = System.Array.Empty<object>();

    /// <summary>Specific object properties to include.</summary>
    [Parameter]
    public string[]? Property { get; set; }

    /// <summary>Header labels. Defaults to property names.</summary>
    [Parameter]
    public string[]? Header { get; set; }

    /// <summary>Table alignment.</summary>
    [Parameter]
    public PdfAlign Align { get; set; } = PdfAlign.Left;

    /// <summary>Emit the updated document.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var document = PdfCommandUtilities.ResolveDocument(this, Document, ParameterSetName, ParameterSetDocument);
        var rowArrayInput = InputObject.All(item => item is IEnumerable && item is not string && item is not IDictionary);
        string[][] rows = rowArrayInput
            ? PdfCommandUtilities.ConvertDataRows(InputObject)
            : InputObject.Length == 1 && InputObject[0] is IEnumerable enumerable && InputObject[0] is not string && InputObject[0] is not IDictionary
            ? PdfCommandUtilities.ConvertDataRows(enumerable)
            : PdfCommandUtilities.ConvertToTableRows(InputObject, Property, Header);

        document.Table(rows, Align);
        if (PassThru.IsPresent)
        {
            WriteObject(document);
        }
    }
}
