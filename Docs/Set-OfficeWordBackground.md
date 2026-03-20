---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficeWordBackground
## SYNOPSIS
Sets the background for a Word document.

## SYNTAX
### Color (Default)
```powershell
Set-OfficeWordBackground [-Color] <string> [-Document <WordDocument>] [-PassThru] [<CommonParameters>]
```

### Image
```powershell
Set-OfficeWordBackground [-ImagePath] <string> [-Document <WordDocument>] [-Width <double>] [-Height <double>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Sets the background for a Word document.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>Set-OfficeWordBackground -Color '#f4f7fb'
```

Sets the document background to the provided hex color.

## PARAMETERS

### -Color
Background color in hex format (#RRGGBB or RRGGBB).

```yaml
Type: String
Parameter Sets: Color
Aliases: None
Possible values: 

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Document
Document to update when provided explicitly.

```yaml
Type: WordDocument
Parameter Sets: Color, Image
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -Height
Optional background image height in pixels.

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

### -ImagePath
Path to the background image.

```yaml
Type: String
Parameter Sets: Image
Aliases: None
Possible values: 

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -PassThru
Emit the updated document.

```yaml
Type: SwitchParameter
Parameter Sets: Color, Image
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Width
Optional background image width in pixels.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Word.WordDocument`

## OUTPUTS

- `OfficeIMO.Word.WordDocument`

## RELATED LINKS

- None

