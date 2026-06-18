---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeVisioStencilShape
## SYNOPSIS
Adds a stencil shape to the current Visio page.

## SYNTAX
### CatalogName (Default)
```powershell
Add-OfficeVisioStencilShape [-Stencil] <string> [[-Text] <string>] [-Page <VisioPage>] [-Catalog <string>] [-Key <string>] [-X <double>] [-Y <double>] [-Width <double>] [-Height <double>] [-ShapeName <string>] [-NameU <string>] [-FillColor <string>] [-LineColor <string>] [-LineWeight <double>] [-LinePattern <int>] [-FillPattern <int>] [-Angle <double>] [<CommonParameters>]
```

### CatalogObject
```powershell
Add-OfficeVisioStencilShape [-Stencil] <string> [[-Text] <string>] -CatalogObject <VisioStencilCatalog> [-Page <VisioPage>] [-Key <string>] [-X <double>] [-Y <double>] [-Width <double>] [-Height <double>] [-ShapeName <string>] [-NameU <string>] [-FillColor <string>] [-LineColor <string>] [-LineWeight <double>] [-LinePattern <int>] [-FillPattern <int>] [-Angle <double>] [<CommonParameters>]
```

### BuiltIn
```powershell
Add-OfficeVisioStencilShape [-Stencil] <string> [[-Text] <string>] [-Page <VisioPage>] [-BuiltIn <OfficeVisioBuiltInStencilCatalog>] [-Key <string>] [-X <double>] [-Y <double>] [-Width <double>] [-Height <double>] [-ShapeName <string>] [-NameU <string>] [-FillColor <string>] [-LineColor <string>] [-LineWeight <double>] [-LinePattern <int>] [-FillPattern <int>] [-Angle <double>] [<CommonParameters>]
```

## DESCRIPTION
Adds a stencil shape to the current Visio page.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> New-OfficeVisio -Path .\StencilFlow.vsdx -UseMastersByDefault {
    Import-OfficeVisioStencil -BuiltIn Flowchart -Name Flow -Default
    VisioStencil -Catalog Flow -Stencil process -Key intake -Text 'Intake' -X 1.5 -Y 4
}
```

Registers a built-in catalog and places a stencil shape on the active page.

## PARAMETERS

### -Angle
Shape angle in radians.

```yaml
Type: Nullable`1
Parameter Sets: CatalogName, CatalogObject, BuiltIn
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -BuiltIn
Built-in OfficeIMO stencil catalog containing the shape.

```yaml
Type: OfficeVisioBuiltInStencilCatalog
Parameter Sets: BuiltIn
Aliases: None
Possible values: All, BasicShapes, Flowchart, BlockDiagram, Architecture, Network, Infrastructure, Cloud, SecurityIdentity, ContainersKubernetes, DataPlatform, CollaborationBusiness, Swimlane, OrgChart, Timeline, Sequence

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Catalog
Catalog previously registered in the active Visio DSL scope.

```yaml
Type: String
Parameter Sets: CatalogName
Aliases: CatalogName
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -CatalogObject
Catalog object containing the stencil shape.

```yaml
Type: VisioStencilCatalog
Parameter Sets: CatalogObject
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -FillColor
Fill color name or hex value.

```yaml
Type: String
Parameter Sets: CatalogName, CatalogObject, BuiltIn
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -FillPattern
Fill pattern.

```yaml
Type: Nullable`1
Parameter Sets: CatalogName, CatalogObject, BuiltIn
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Height
Optional shape height. Omit to use the stencil default height.

```yaml
Type: Nullable`1
Parameter Sets: CatalogName, CatalogObject, BuiltIn
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Key
DSL key used by connector commands.

```yaml
Type: String
Parameter Sets: CatalogName, CatalogObject, BuiltIn
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -LineColor
Line color name or hex value.

```yaml
Type: String
Parameter Sets: CatalogName, CatalogObject, BuiltIn
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -LinePattern
Line pattern.

```yaml
Type: Nullable`1
Parameter Sets: CatalogName, CatalogObject, BuiltIn
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -LineWeight
Line weight.

```yaml
Type: Nullable`1
Parameter Sets: CatalogName, CatalogObject, BuiltIn
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -NameU
Optional universal shape name.

```yaml
Type: String
Parameter Sets: CatalogName, CatalogObject, BuiltIn
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Page
Target page. Optional inside VisioPage or New-OfficeVisio.

```yaml
Type: VisioPage
Parameter Sets: CatalogName, CatalogObject, BuiltIn
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -ShapeName
Optional shape name.

```yaml
Type: String
Parameter Sets: CatalogName, CatalogObject, BuiltIn
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Stencil
Stencil id, name, master name, keyword, alias, or tag.

```yaml
Type: String
Parameter Sets: CatalogName, CatalogObject, BuiltIn
Aliases: Name
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Text
Text placed inside the shape. Omit to use the stencil display name.

```yaml
Type: String
Parameter Sets: CatalogName, CatalogObject, BuiltIn
Aliases: None
Possible values:

Required: False
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Width
Optional shape width. Omit to use the stencil default width.

```yaml
Type: Nullable`1
Parameter Sets: CatalogName, CatalogObject, BuiltIn
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -X
X coordinate of the stencil shape center.

```yaml
Type: Double
Parameter Sets: CatalogName, CatalogObject, BuiltIn
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Y
Y coordinate of the stencil shape center.

```yaml
Type: Double
Parameter Sets: CatalogName, CatalogObject, BuiltIn
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

- `OfficeIMO.Visio.VisioPage`

## OUTPUTS

- `OfficeIMO.Visio.VisioShape`

## RELATED LINKS

- None
