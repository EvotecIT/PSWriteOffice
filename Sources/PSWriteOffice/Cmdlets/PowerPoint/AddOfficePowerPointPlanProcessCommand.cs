using System.Management.Automation;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Adds a semantic process/timeline slide to a PowerPoint deck plan.</summary>
[Cmdlet(VerbsCommon.Add, "OfficePowerPointPlanProcess")]
[Alias("PptPlanProcess")]
[OutputType(typeof(PowerPointDeckPlan))]
public sealed class AddOfficePowerPointPlanProcessCommand : PSCmdlet
{
    /// <summary>Plan to update. Optional inside New-OfficePowerPointDeckPlan.</summary>
    [Parameter(ValueFromPipeline = true)]
    public PowerPointDeckPlan? Plan { get; set; }

    /// <summary>Slide title.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Title { get; set; } = string.Empty;

    /// <summary>Optional slide subtitle.</summary>
    [Parameter]
    public string? Subtitle { get; set; }

    /// <summary>Objects with Title and Body/Description/Text properties.</summary>
    [Parameter(Mandatory = true)]
    public object[] Steps { get; set; } = System.Array.Empty<object>();

    /// <summary>Stable seed for deterministic visual selection.</summary>
    [Parameter]
    public string? Seed { get; set; }

    /// <summary>Emit the updated plan.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var plan = Plan ?? PowerPointDeckPlanDslContext.Require(this).Plan;
        plan.AddProcess(Title, Subtitle, PowerPointDesignerDataMapper.ToProcessSteps(Steps), Seed);
        if (PassThru.IsPresent)
        {
            WriteObject(plan);
        }
    }
}
