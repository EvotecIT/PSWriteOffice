using System.Management.Automation;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Creates a semantic PowerPoint deck plan for designer rendering.</summary>
[Cmdlet(VerbsCommon.New, "OfficePowerPointDeckPlan")]
[Alias("PptDeckPlan")]
[OutputType(typeof(PowerPointDeckPlan))]
public sealed class NewOfficePowerPointDeckPlanCommand : PSCmdlet
{
    /// <summary>Nested deck-plan DSL content.</summary>
    [Parameter(Position = 0)]
    public ScriptBlock? Content { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var plan = new PowerPointDeckPlan();
        if (Content != null)
        {
            using (PowerPointDeckPlanDslContext.Enter(plan))
            {
                Content.InvokeReturnAsIs();
            }
        }

        WriteObject(plan);
    }
}
