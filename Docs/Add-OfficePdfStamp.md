---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficePdfStamp
## SYNOPSIS
Adds a text or image stamp to an existing PDF.

## SYNTAX
### Text (Default)
```powershell
Add-OfficePdfStamp -Path <string> -OutputPath <string> -Text <string> [-PageRange <string>] [-X <double>] [-Y <double>] [-FontSize <double>] [-Color <string>] [-Rotation <double>] [-Watermark] [<CommonParameters>]
```

### Image
```powershell
Add-OfficePdfStamp -Path <string> -OutputPath <string> -Image <string> [-PageRange <string>] [-X <double>] [-Y <double>] [-Width <double>] [-Height <double>] [-Rotation <double>] [-Watermark] [<CommonParameters>]
```

## DESCRIPTION
Stamps are existing-PDF operations. Use text stamps for review labels and image stamps for logos or approval marks.
Use -Watermark when the stamp should be placed behind existing page content.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Add-OfficePdfStamp -Path .\Examples\Documents\Report.pdf -OutputPath .\Examples\Documents\Stamped.pdf -Text 'REVIEWED' -Color '#0F766E' -FontSize 24 -Rotation 12 -PageRange '1-2'
            Get-OfficePdfPreflight -Path .\Examples\Documents\Stamped.pdf
```

Adds a text stamp to the first two pages and preflights the result.

### EXAMPLE 2
```powershell
PS> $logo = '.\Tests\Assets\CellImage.png'
Add-OfficePdfStamp -Path .\Examples\Documents\Report.pdf -OutputPath .\Examples\Documents\Watermarked.pdf -Image $logo -Width 160 -Watermark
```

Adds a logo behind existing content as a watermark.

## PARAMETERS

### -Color
Text color in #RRGGBB format.

```yaml
Type: String
Parameter Sets: Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -FontSize
Font size for text stamps.

```yaml
Type: Double
Parameter Sets: Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Height
Rendered image height in PDF points.

```yaml
Type: Nullable`1
Parameter Sets: Image
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Image
Image path to stamp.

```yaml
Type: String
Parameter Sets: Image
Aliases: ImagePath
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -OutputPath
Output PDF path.

```yaml
Type: String
Parameter Sets: Text, Image
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -PageRange
Stamp selected pages, for example 1-3,5. Omit to stamp every page.

```yaml
Type: String
Parameter Sets: Text, Image
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Path
Input PDF path.

```yaml
Type: String
Parameter Sets: Text, Image
Aliases: FilePath
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Rotation
Rotation in degrees.

```yaml
Type: Double
Parameter Sets: Text, Image
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Text
Text to stamp.

```yaml
Type: String
Parameter Sets: Text
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Watermark
Place the stamp behind existing content as a watermark.

```yaml
Type: SwitchParameter
Parameter Sets: Text, Image
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Width
Rendered image width in PDF points.

```yaml
Type: Nullable`1
Parameter Sets: Image
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -X
X coordinate in PDF points.

```yaml
Type: Nullable`1
Parameter Sets: Text, Image
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Y
Y coordinate in PDF points.

```yaml
Type: Nullable`1
Parameter Sets: Text, Image
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

- `None`

## OUTPUTS

- `System.IO.FileInfo`

## RELATED LINKS

- None
