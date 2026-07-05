param(
    [switch] $Open
)

Import-Module PSWriteOffice -ErrorAction Stop

$documents = Join-Path $PSScriptRoot '..\Documents'
New-Item -Path $documents -ItemType Directory -Force | Out-Null

$path = Join-Path $documents 'Showcase-PowerPoint-ServiceBrief.pptx'

$process = @(
    [pscustomobject]@{ Title = 'Assess'; Body = 'Map services, owners, risk, and reporting expectations.' }
    [pscustomobject]@{ Title = 'Design'; Body = 'Choose workbook, report, and deck outputs that match the audience.' }
    [pscustomobject]@{ Title = 'Automate'; Body = 'Generate Office files from repeatable PowerShell objects.' }
    [pscustomobject]@{ Title = 'Publish'; Body = 'Attach visuals and explain the workflow in reusable blog posts.' }
)

$cards = @(
    [pscustomobject]@{ Title = 'Word'; Items = 'TOC|sections|tables|charts|approvals'; AccentColor = '#2F80ED' }
    [pscustomobject]@{ Title = 'Excel'; Items = 'dashboard|pivots|sparklines|validation|links'; AccentColor = '#219653' }
    [pscustomobject]@{ Title = 'PowerPoint'; Items = 'designer plans|process|cards|coverage|notes'; AccentColor = '#9B51E0' }
    [pscustomobject]@{ Title = 'Blog'; Items = 'screenshots|code|generated covers|artifact links'; AccentColor = '#F2994A' }
)

$coverage = @(
    [pscustomobject]@{ Name = 'Engine'; X = 0.22; Y = 0.42; Detail = 'OfficeIMO owns Open XML behavior.' }
    [pscustomobject]@{ Name = 'PowerShell'; X = 0.48; Y = 0.34; Detail = 'PSWriteOffice owns scripting ergonomics.' }
    [pscustomobject]@{ Name = 'Examples'; X = 0.68; Y = 0.58; Detail = 'Showcase scripts prove the surface.' }
    [pscustomobject]@{ Name = 'Website'; X = 0.82; Y = 0.36; Detail = 'Blog posts turn artifacts into adoption.' }
)

$capabilities = @(
    [pscustomobject]@{ Heading = 'Readable by humans'; Body = 'Outputs should look like business artifacts, not raw exports.'; Items = 'visual hierarchy|navigation|metadata' }
    [pscustomobject]@{ Heading = 'Useful to scripts'; Body = 'Generated files should be inspectable and testable.'; Items = 'summaries|parts|deterministic paths' }
    [pscustomobject]@{ Heading = 'Fast enough to reuse'; Body = 'Examples should be smoke-test friendly and avoid desktop Office for generation.'; Items = 'Open XML|parallel engine work|small fixtures' }
)

$caseStudy = @(
    [pscustomobject]@{ Heading = 'Problem'; Body = 'Basic examples hide how much OfficeIMO can already produce.' }
    [pscustomobject]@{ Heading = 'Approach'; Body = 'Expose semantic PowerPoint plans and richer Office examples through PSWriteOffice.' }
    [pscustomobject]@{ Heading = 'Outcome'; Body = 'A test-drive deck that remains editable and visually credible.' }
)

$metrics = @(
    [pscustomobject]@{ Value = '3'; Label = 'flagship products' }
    [pscustomobject]@{ Value = '1'; Label = 'shared showcase plan' }
    [pscustomobject]@{ Value = '0'; Label = 'desktop Office dependency' }
)

$chartRows = @(
    [pscustomobject]@{ Product = 'Word'; Coverage = 82; Polish = 72 }
    [pscustomobject]@{ Product = 'Excel'; Coverage = 91; Polish = 83 }
    [pscustomobject]@{ Product = 'PowerPoint'; Coverage = 76; Polish = 68 }
)

$tableRows = @(
    [pscustomobject]@{ Area = 'Designer bridge'; Status = 'Added'; Next = 'Add more variant knobs' }
    [pscustomobject]@{ Area = 'Showcase deck'; Status = 'Added'; Next = 'Export screenshots' }
    [pscustomobject]@{ Area = 'Blog post'; Status = 'Planned'; Next = 'Write after visuals' }
)

$plan = PptDeckPlan {
    PptPlanSection -Title 'PSWriteOffice Showcase' -Subtitle 'Beautiful, useful Office artifacts from PowerShell' -Seed 'showcase-cover'
    PptPlanProcess -Title 'From objects to publishable artifacts' -Subtitle 'A repeatable path for examples and blog posts' -Steps $process -Seed 'delivery-process'
    PptPlanCardGrid -Title 'Product surfaces' -Subtitle 'Each product should show a real workflow, not only primitives.' -Cards $cards -Seed 'product-cards'
    PptPlanCoverage -Title 'Where the work belongs' -Subtitle 'Engine, wrapper, examples, and website stay distinct.' -Locations $coverage -Seed 'coverage-map'
    PptPlanCapability -Title 'Quality bar' -Subtitle 'The showcase should be practical enough to copy and attractive enough to publish.' -Sections $capabilities -Seed 'quality-bar'
    PptPlanCaseStudy -Title 'PowerPoint designer bridge' -Sections $caseStudy -Metrics $metrics -Seed 'designer-case-study'
}

New-OfficePowerPoint -Path $path {
    PptSlideSize -Preset Screen16x9 | Out-Null
    PptDesignerDeck -Plan $plan -AccentColor '#008C95' -Seed 'pswriteoffice-showcase' -Purpose 'technical service brief' -Name 'PSWriteOffice Showcase' -FooterLeft 'PSWriteOffice' -FooterRight 'OfficeIMO designer' -CreativeDirectionPack TechnicalMap -LayoutStrategy ContentFirst | Out-Null

    PptSlide {
        PptTitle -Title 'Coverage and polish scorecard' | Out-Null
        PptChart -Type ClusteredColumn -Data $chartRows -CategoryProperty Product -SeriesProperty Coverage, Polish -Title 'Current Surface vs Polish Target' -X 58 -Y 118 -Width 610 -Height 265 | Out-Null
        PptNotes -Text 'Use this slide as the bridge between the designer slides and the concrete backlog.'
    } | Out-Null

    PptSlide {
        PptTitle -Title 'Immediate implementation path' | Out-Null
        PptTable -Data $tableRows -X 64 -Y 132 -Width 590 -Height 210 | Out-Null
        PptNotes -Text 'Close with the next concrete pull request slices: visual screenshots, blog drafts, and richer wrappers.'
    } | Out-Null
} -Open:$Open

$ppt = Get-OfficePowerPoint -FilePath $path
try {
    Add-OfficePowerPointSection -Presentation $ppt -Name 'Designer story' -StartSlideIndex 0 | Out-Null
    Add-OfficePowerPointSection -Presentation $ppt -Name 'Evidence appendix' -StartSlideIndex 6 | Out-Null
    Get-OfficePowerPointSlide -Presentation $ppt -Index 0 | Set-OfficePowerPointSlideTransition -Transition Fade | Out-Null
    Get-OfficePowerPointSlide -Presentation $ppt -Index 6 | Set-OfficePowerPointSlideTransition -Transition PushLeft | Out-Null
    Save-OfficePowerPoint -Presentation $ppt
}
finally {
    $ppt.Dispose()
}

$summary = Get-OfficePowerPoint -FilePath $path | Get-OfficePowerPointSlideSummary
$summary | Format-Table Index, Title, ShapeCount, TextBoxCount, ChartCount, TableCount, HasNotes

Write-Host "Presentation saved to $path"
