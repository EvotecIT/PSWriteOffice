---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Export-OfficeVisioStencilPreviewGallery
## SYNOPSIS
Exports preview artwork from a Visio stencil package into a browsable HTML gallery.

## SYNTAX
### __AllParameterSets
```powershell
Export-OfficeVisioStencilPreviewGallery [-Path] <string> [-OutputDirectory] <string> [-Title <string>] [-MasterName <string[]>] [-IncludeUnsupportedMasters] [-NoLearnMasterDimensions] [-PreviewDirectoryName <string>] [-IndexFileName <string>] [-NoIndex] [-NoThumbnails] [-ThumbnailDirectoryName <string>] [-ThumbnailWidth <int>] [-ThumbnailHeight <int>] [-DefaultWidth <double>] [-DefaultHeight <double>] [<CommonParameters>]
```

## DESCRIPTION
Exports preview artwork from a Visio stencil package into a browsable HTML gallery.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $gallery = Export-OfficeVisioStencilPreviewGallery -Path .\MyShapes.vssx -OutputDirectory .\StencilGallery -Title 'Custom stencil previews'
$gallery | Select-Object PackagePath, IndexPath, BrowserRenderableCount, ThumbnailCount
```

Extracts preview artwork from package-backed masters and writes preview files plus an HTML index.

## PARAMETERS

### -DefaultHeight
Default height for package-backed stencils when dimensions cannot be learned.

```yaml
Type: Double
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -DefaultWidth
Default width for package-backed stencils when dimensions cannot be learned.

```yaml
Type: Double
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -IncludeUnsupportedMasters
Include unsupported package masters when looking for preview artwork.

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

### -IndexFileName
Generated HTML index file name.

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

### -MasterName
Optional master filters for package-backed catalogs.

```yaml
Type: String[]
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -NoIndex
Do not write an HTML index file.

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

### -NoLearnMasterDimensions
Skip reading master dimensions from package master parts.

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

### -NoThumbnails
Do not write browser-renderable thumbnail wrappers.

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

### -OutputDirectory
Directory that receives preview payloads, thumbnails, and the HTML index.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Path
Visio package path, such as .vsdx, .vssx, or .vstx.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: FilePath, LiteralPath
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -PreviewDirectoryName
Subdirectory that receives extracted preview payload files.

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

### -ThumbnailDirectoryName
Subdirectory that receives generated thumbnail wrappers.

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

### -ThumbnailHeight
Generated thumbnail height in pixels.

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

### -ThumbnailWidth
Generated thumbnail width in pixels.

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

### -Title
Gallery title. When omitted, a title is derived from the package name.

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

- `None`

## OUTPUTS

- `OfficeIMO.Visio.Stencils.VisioStencilPreviewGallery`

## RELATED LINKS

- None
