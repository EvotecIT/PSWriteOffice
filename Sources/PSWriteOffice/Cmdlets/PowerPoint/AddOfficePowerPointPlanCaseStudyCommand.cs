using System.Management.Automation;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Adds a semantic case-study slide to a PowerPoint deck plan.</summary>
[Cmdlet(VerbsCommon.Add, "OfficePowerPointPlanCaseStudy")]
[Alias("PptPlanCaseStudy")]
[OutputType(typeof(PowerPointDeckPlan))]
public sealed class AddOfficePowerPointPlanCaseStudyCommand : PSCmdlet
{
    /// <summary>Plan to update. Optional inside New-OfficePowerPointDeckPlan.</summary>
    [Parameter(ValueFromPipeline = true)]
    public PowerPointDeckPlan? Plan { get; set; }

    /// <summary>Slide title.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Title { get; set; } = string.Empty;

    /// <summary>Objects with Heading/Title and Body/Description/Text properties.</summary>
    [Parameter(Mandatory = true)]
    public object[] Sections { get; set; } = System.Array.Empty<object>();

    /// <summary>Objects with Value and Label/Name/Title properties.</summary>
    [Parameter]
    public object[]? Metrics { get; set; }

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
        plan.AddCaseStudy(
            Title,
            PowerPointDesignerDataMapper.ToCaseStudySections(Sections),
            PowerPointDesignerDataMapper.ToMetrics(Metrics),
            Seed);
        if (PassThru.IsPresent)
        {
            WriteObject(plan);
        }
    }
}
