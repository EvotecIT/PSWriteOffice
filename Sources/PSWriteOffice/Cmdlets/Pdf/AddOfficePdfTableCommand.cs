using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;
using PSWriteOffice.Services.Table;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Adds a table to a PDF document.</summary>
/// <example>
///   <summary>Add object data as a PDF table.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$services = @(
///     [pscustomobject]@{ Name = 'Directory'; Status = 'Healthy'; Incidents = 0 }
///     [pscustomobject]@{ Name = 'Mail'; Status = 'Watch'; Incidents = 2 }
/// )
/// New-OfficePdf -Path .\Examples\Documents\PdfTable.pdf {
///     Add-OfficePdfHeading -Text 'Service status'
///     Add-OfficePdfTable -InputObject $services -Property Name,Status,Incidents -Header 'Service','Status','Incidents'
/// }</code>
///   <para>Converts PowerShell objects into a table using selected properties and friendly headers.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficePdfTable", DefaultParameterSetName = ParameterSetContext)]
[Alias("PdfTable")]
[OutputType(typeof(PdfDocument))]
public sealed class AddOfficePdfTableCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";
    private const string ParameterSetPipelineDocument = "PipelineDocument";
    private readonly List<object?> _items = new();

    /// <summary>PDF document to update outside the DSL context.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetDocument)]
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetPipelineDocument)]
    public PdfDocument Document { get; set; } = null!;

    /// <summary>Objects or row arrays to render as a table.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = ParameterSetContext)]
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPipelineDocument)]
    public object? InputObject { get; set; }

    /// <summary>Specific object properties to include.</summary>
    [Parameter]
    public string[]? Property { get; set; }

    /// <summary>Header labels. Defaults to property names.</summary>
    [Parameter]
    public string[]? Header { get; set; }

    /// <summary>Projection to apply before writing the table.</summary>
    [Parameter]
    public OfficeTableView View { get; set; } = OfficeTableView.Normal;

    /// <summary>Table alignment.</summary>
    [Parameter]
    public PdfAlign Align { get; set; } = PdfAlign.Left;

    /// <summary>Emit the updated document.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (ParameterSetName == ParameterSetPipelineDocument)
        {
            RenderTable(Document, BuildRows(InputObject));
            if (PassThru.IsPresent)
            {
                WriteObject(Document);
            }

            return;
        }

        TableInputCollector.AddInput(_items, InputObject);
    }

    /// <inheritdoc />
    protected override void EndProcessing()
    {
        if (ParameterSetName == ParameterSetPipelineDocument)
        {
            return;
        }

        var document = PdfCommandUtilities.ResolveDocument(this, Document, ParameterSetName, ParameterSetDocument);
        RenderTable(document, TableInputCollector.RequireRows(_items, nameof(InputObject)));
        if (PassThru.IsPresent)
        {
            WriteObject(document);
        }
    }

    private void RenderTable(PdfDocument document, object[] inputRows)
    {
        var projectedRows = TableViewProjection.Project(inputRows, View);
        var rowArrayInput = projectedRows.All(item => item is IEnumerable && item is not string && item is not IDictionary);
        string[][] rows = rowArrayInput
            ? PdfCommandUtilities.ConvertDataRows(projectedRows, Header)
            : projectedRows.Length == 1 && projectedRows[0] is IEnumerable enumerable && projectedRows[0] is not string && projectedRows[0] is not IDictionary
            ? PdfCommandUtilities.ConvertDataRows(enumerable, Header)
            : PdfCommandUtilities.ConvertToTableRows(projectedRows, Property, Header);

        document.Table(rows, Align);
    }

    private static object[] BuildRows(object? value)
    {
        var items = new List<object?>();
        TableInputCollector.AddInput(items, value);
        return TableInputCollector.RequireRows(items, nameof(InputObject));
    }
}
