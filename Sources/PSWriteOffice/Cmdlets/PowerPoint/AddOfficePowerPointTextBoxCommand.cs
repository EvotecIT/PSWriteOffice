using System;
using System.Management.Automation;
using ShapeCrawler;

namespace PSWriteOffice.Cmdlets.PowerPoint;

[Cmdlet(VerbsCommon.Add, "OfficePowerPointTextBox")]
public class AddOfficePowerPointTextBoxCommand : PSCmdlet
{
    [Parameter(Mandatory = true, ValueFromPipeline = true)]
    public ISlide Slide { get; set; } = null!;

    [Parameter(Mandatory = true)]
    public string Text { get; set; } = null!;

    [Parameter]
    public int X { get; set; } = 50;

    [Parameter]
    public int Y { get; set; } = 50;

    [Parameter]
    public int Width { get; set; } = 200;

    [Parameter]
    public int Height { get; set; } = 50;

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
