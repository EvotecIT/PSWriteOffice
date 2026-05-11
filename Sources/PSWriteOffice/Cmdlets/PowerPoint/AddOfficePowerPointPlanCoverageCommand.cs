using System.Management.Automation;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Adds a semantic coverage/location slide to a PowerPoint deck plan.</summary>
[Cmdlet(VerbsCommon.Add, "OfficePowerPointPlanCoverage")]
[Alias("PptPlanCoverage")]
[OutputType(typeof(PowerPointDeckPlan))]
public sealed class AddOfficePowerPointPlanCoverageCommand : PSCmdlet
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

    /// <summary>Objects with Name, X, Y, and optional Detail properties. X/Y are normalized 0..1 positions.</summary>
    [Parameter(Mandatory = true)]
    public object[] Locations { get; set; } = System.Array.Empty<object>();

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
        plan.AddCoverage(Title, Subtitle, PowerPointDesignerDataMapper.ToLocations(Locations), Seed);
        if (PassThru.IsPresent)
        {
            WriteObject(plan);
        }
    }
}
