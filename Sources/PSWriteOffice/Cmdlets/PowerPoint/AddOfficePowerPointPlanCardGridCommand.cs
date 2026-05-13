using System.Management.Automation;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Adds a semantic card-grid slide to a PowerPoint deck plan.</summary>
[Cmdlet(VerbsCommon.Add, "OfficePowerPointPlanCardGrid")]
[Alias("PptPlanCardGrid")]
[OutputType(typeof(PowerPointDeckPlan))]
public sealed class AddOfficePowerPointPlanCardGridCommand : PSCmdlet
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

    /// <summary>Objects with Title/Name plus optional Items/Bullets/Details and AccentColor properties.</summary>
    [Parameter(Mandatory = true)]
    public object[] Cards { get; set; } = System.Array.Empty<object>();

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
        plan.AddCardGrid(Title, Subtitle, PowerPointDesignerDataMapper.ToCards(Cards), Seed);
        if (PassThru.IsPresent)
        {
            WriteObject(plan);
        }
    }
}
