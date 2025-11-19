using System;
using System.Management.Automation;
using ShapeCrawler;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Adds a text box to a slide.</summary>
/// <para>Creates a rectangle at the requested coordinates and assigns the supplied text.</para>
/// <example>
///   <summary>Insert a caption.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Add-OfficePowerPointTextBox -Slide $slide -Text 'Quarterly Overview' -X 80 -Y 150</code>
///   <para>Places a text box mid-slide with the provided caption.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficePowerPointTextBox")]
public class AddOfficePowerPointTextBoxCommand : PSCmdlet
{
    /// <summary>Target slide that will receive the text box.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true)]
    public ISlide Slide { get; set; } = null!;

    /// <summary>Text to render inside the box.</summary>
    [Parameter(Mandatory = true)]
    public string Text { get; set; } = null!;

    /// <summary>Left offset (in points) from the slide origin.</summary>
    [Parameter]
    public int X { get; set; } = 50;

    /// <summary>Top offset (in points) from the slide origin.</summary>
    [Parameter]
    public int Y { get; set; } = 50;

    /// <summary>Box width in points.</summary>
    [Parameter]
    public int Width { get; set; } = 200;

    /// <summary>Box height in points.</summary>
    [Parameter]
    public int Height { get; set; } = 50;

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        try
        {
            var shapes = Slide.Shapes;
            shapes.AddShape(X, Y, Width, Height);
            var shape = shapes[shapes.Count - 1];
            shape.TextBox!.SetText(Text);
            WriteObject(shape);
        }
        catch (Exception ex)
        {
            WriteError(new ErrorRecord(ex, "PowerPointAddTextBoxFailed", ErrorCategory.InvalidOperation, Slide));
        }
    }
}
