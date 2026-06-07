---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficePowerPointDesignerDeck
## SYNOPSIS
Renders a semantic deck plan through OfficeIMO PowerPoint designer helpers.

## SYNTAX
### __AllParameterSets
```powershell
Add-OfficePowerPointDesignerDeck -Plan <PowerPointDeckPlan> [-Presentation <PowerPointPresentation>] [-AccentColor <string>] [-Seed <string>] [-Purpose <string>] [-Name <string>] [-Eyebrow <string>] [-FooterLeft <string>] [-FooterRight <string>] [-CreativeDirectionPack <string>] [-LayoutStrategy <string>] [-AlternativeCount <int>] [-NoApplyTheme] [-Preview] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Renders a semantic deck plan through OfficeIMO PowerPoint designer helpers.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $plan = New-OfficePowerPointDeckPlan {
    Add-OfficePowerPointPlanSection -Title 'Service Review'
    Add-OfficePowerPointPlanCardGrid -Title 'Current signals' -Cards @(
      @{ Title = 'Availability'; Items = @('Healthy', 'No critical incidents') }
      @{ Title = 'Risk'; Items = @('One dependency on watch') }
    )
}
New-OfficePowerPoint -Path .\Examples\Documents\DesignerDeck.pptx {
    Add-OfficePowerPointDesignerDeck -Plan $plan -AccentColor '#0F766E' -Purpose 'monthly service brief'
}
```

Uses OfficeIMO design selection to turn semantic content into slides.

## PARAMETERS

### -AccentColor
Brand accent color used to derive the deck palette.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -AlternativeCount
Design alternative count to consider. 0 uses OfficeIMO defaults.

```yaml
Type: Int32
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -CreativeDirectionPack
Creative direction pack name, such as Boardroom, FieldProof, TechnicalMap, or QuietAppendix.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Eyebrow
Default slide eyebrow.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -FooterLeft
Left footer text.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -FooterRight
Right footer text.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -LayoutStrategy
Auto layout strategy, such as ContentFirst, DesignFirst, Compact, or VisualFirst.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Name
Deck theme name.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -NoApplyTheme
Do not automatically apply the design theme to the presentation.

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -PassThru
Emit rendered slides instead of the render summary.

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Plan
Deck plan to render.

```yaml
Type: PowerPointDeckPlan
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Presentation
Presentation to update. Optional inside New-OfficePowerPoint.

```yaml
Type: PowerPointPresentation
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -Preview
Preview resolved slides without rendering them.

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Purpose
Plain-language purpose used to select a built-in design recipe.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Seed
Stable seed used for deterministic design choices.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.PowerPoint.PowerPointPresentation`

## OUTPUTS

- `OfficeIMO.PowerPoint.PowerPointDeckPlanSlideRenderSummary
OfficeIMO.PowerPoint.PowerPointSlide`

## RELATED LINKS

- None
