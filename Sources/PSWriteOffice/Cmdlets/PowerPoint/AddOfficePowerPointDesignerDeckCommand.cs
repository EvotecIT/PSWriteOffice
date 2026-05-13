using System;
using System.Management.Automation;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Services;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Renders a semantic deck plan through OfficeIMO PowerPoint designer helpers.</summary>
[Cmdlet(VerbsCommon.Add, "OfficePowerPointDesignerDeck")]
[Alias("PptDesignerDeck")]
[OutputType(typeof(PowerPointDeckPlanSlideRenderSummary), typeof(PowerPointSlide))]
public sealed class AddOfficePowerPointDesignerDeckCommand : PSCmdlet
{
    /// <summary>Presentation to update. Optional inside New-OfficePowerPoint.</summary>
    [Parameter(ValueFromPipeline = true)]
    public PowerPointPresentation? Presentation { get; set; }

    /// <summary>Deck plan to render.</summary>
    [Parameter(Mandatory = true)]
    public PowerPointDeckPlan Plan { get; set; } = null!;

    /// <summary>Brand accent color used to derive the deck palette.</summary>
    [Parameter]
    public string AccentColor { get; set; } = "#008C95";

    /// <summary>Stable seed used for deterministic design choices.</summary>
    [Parameter]
    public string Seed { get; set; } = "pswriteoffice";

    /// <summary>Plain-language purpose used to select a built-in design recipe.</summary>
    [Parameter]
    public string Purpose { get; set; } = "technical service brief";

    /// <summary>Deck theme name.</summary>
    [Parameter]
    public string? Name { get; set; }

    /// <summary>Default slide eyebrow.</summary>
    [Parameter]
    public string? Eyebrow { get; set; }

    /// <summary>Left footer text.</summary>
    [Parameter]
    public string? FooterLeft { get; set; }

    /// <summary>Right footer text.</summary>
    [Parameter]
    public string? FooterRight { get; set; }

    /// <summary>Creative direction pack name, such as Boardroom, FieldProof, TechnicalMap, or QuietAppendix.</summary>
    [Parameter]
    public string? CreativeDirectionPack { get; set; }

    /// <summary>Auto layout strategy, such as ContentFirst, DesignFirst, Compact, or VisualFirst.</summary>
    [Parameter]
    public string? LayoutStrategy { get; set; }

    /// <summary>Design alternative count to consider. 0 uses OfficeIMO defaults.</summary>
    [Parameter]
    public int AlternativeCount { get; set; } = 3;

    /// <summary>Do not automatically apply the design theme to the presentation.</summary>
    [Parameter]
    public SwitchParameter NoApplyTheme { get; set; }

    /// <summary>Preview resolved slides without rendering them.</summary>
    [Parameter]
    public SwitchParameter Preview { get; set; }

    /// <summary>Emit rendered slides instead of the render summary.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (Plan == null)
        {
            throw new PSArgumentException("Provide a deck plan.", nameof(Plan));
        }

        if (AlternativeCount < 0)
        {
            throw new PSArgumentException("AlternativeCount must be zero or greater.", nameof(AlternativeCount));
        }

        var brief = PowerPointDesignBrief
            .FromBrand(AccentColor, Seed, Purpose)
            .WithIdentity(Name, Eyebrow, FooterLeft, FooterRight);

        if (!string.IsNullOrWhiteSpace(CreativeDirectionPack))
        {
            if (!OpenXmlValueParser.TryParse(CreativeDirectionPack, out PowerPointCreativeDirectionPack pack))
            {
                throw new PSArgumentException($"Unknown CreativeDirectionPack '{CreativeDirectionPack}'.", nameof(CreativeDirectionPack));
            }

            brief.WithCreativeDirectionPack(pack);
        }

        if (!string.IsNullOrWhiteSpace(LayoutStrategy))
        {
            if (!OpenXmlValueParser.TryParse(LayoutStrategy, out PowerPointAutoLayoutStrategy strategy))
            {
                throw new PSArgumentException($"Unknown LayoutStrategy '{LayoutStrategy}'.", nameof(LayoutStrategy));
            }

            brief.WithLayoutStrategy(strategy);
        }

        if (Preview.IsPresent)
        {
            WriteObject(brief.DescribeDeckPlan(Plan, 0), enumerateCollection: true);
            return;
        }

        var presentation = Presentation ?? PowerPointDslContext.Require(this).Presentation;
        var composer = presentation.UseDesigner(
            brief,
            Plan,
            AlternativeCount,
            applyTheme: !NoApplyTheme.IsPresent);

        var summary = composer.DescribeSlides(Plan);
        var slides = composer.AddSlides(Plan);

        if (PassThru.IsPresent)
        {
            WriteObject(slides, enumerateCollection: true);
        }
        else
        {
            WriteObject(summary, enumerateCollection: true);
        }
    }
}
