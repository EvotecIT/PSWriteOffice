using System;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using OfficeIMO.Confluence;
using PSWriteOffice.Cmdlets.Confluence;

namespace PSWriteOffice.Tests;

public sealed class ConfluenceCmdletTests
{
    [Fact]
    public void PublishPlanOnlyBuildsAdfRequestWithoutSessionOrNetwork()
    {
        using var runspace = CreateRunspace(
            "Publish-OfficeConfluencePage",
            typeof(PublishOfficeConfluencePageCommand));
        using var powerShell = PowerShell.Create();
        powerShell.Runspace = runspace;
        powerShell
            .AddCommand("Publish-OfficeConfluencePage")
            .AddParameter("SpaceId", "42")
            .AddParameter("Title", "Daily status")
            .AddParameter("Content", "# Ready")
            .AddParameter("PlanOnly");

        var output = powerShell.Invoke();

        Assert.False(powerShell.HadErrors, string.Join(Environment.NewLine, powerShell.Streams.Error));
        var plan = Assert.IsType<ConfluencePageWritePlan>(Assert.Single(output).BaseObject);
        Assert.Equal("POST", plan.Method);
        Assert.Contains("atlas_doc_format", plan.Payload);
        Assert.Contains("Daily status", plan.Payload);
    }

    [Fact]
    public void ManagedSectionCmdletPreservesUnmanagedContent()
    {
        using var runspace = CreateRunspace(
            "Set-OfficeConfluenceManagedSection",
            typeof(SetOfficeConfluenceManagedSectionCommand));
        using var powerShell = PowerShell.Create();
        powerShell.Runspace = runspace;
        powerShell
            .AddCommand("Set-OfficeConfluenceManagedSection")
            .AddParameter("ExistingBody", "<p>owner</p>")
            .AddParameter("SectionId", "daily")
            .AddParameter("Replacement", "<p>generated</p>")
            .AddParameter("AppendIfMissing");

        var output = powerShell.Invoke();

        Assert.False(powerShell.HadErrors, string.Join(Environment.NewLine, powerShell.Streams.Error));
        var result = Assert.IsType<ConfluenceManagedSectionResult>(Assert.Single(output).BaseObject);
        Assert.True(result.WasCreated);
        Assert.Contains("<p>owner</p>", result.UpdatedBody);
        Assert.Contains("<p>generated</p>", result.UpdatedBody);
        Assert.NotEqual(result.OriginalSha256, result.UpdatedSha256);
    }

    [Fact]
    public void LiveWriteCmdletsDeclareShouldProcess()
    {
        AssertShouldProcess(typeof(PublishOfficeConfluencePageCommand));
        AssertShouldProcess(typeof(SendOfficeConfluenceAttachmentCommand));
    }

    private static Runspace CreateRunspace(string name, Type cmdletType)
    {
        var state = InitialSessionState.CreateDefault();
        state.Commands.Add(new SessionStateCmdletEntry(name, cmdletType, helpFileName: null));
        var runspace = RunspaceFactory.CreateRunspace(state);
        runspace.Open();
        return runspace;
    }

    private static void AssertShouldProcess(Type cmdletType)
    {
        var attribute = Assert.Single(cmdletType.GetCustomAttributes(typeof(CmdletAttribute), inherit: false).Cast<CmdletAttribute>());
        Assert.True(attribute.SupportsShouldProcess);
    }
}
