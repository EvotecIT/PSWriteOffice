---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Save-OfficeEmailMailbox
## SYNOPSIS
Saves a native mbox mailbox with output diagnostics.

## SYNTAX
### __AllParameterSets
```powershell
Save-OfficeEmailMailbox [-Path] <string> -Mailbox <EmailMailbox> [-Options <EmailMailboxWriterOptions>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Saves a native mbox mailbox with output diagnostics.

## EXAMPLES

### EXAMPLE 1
```powershell
Save-OfficeEmailMailbox -Mailbox 'Value'
```


## PARAMETERS

### -Mailbox
Mailbox to save.

```yaml
Type: EmailMailbox
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -Options
Optional mailbox variant, envelope, and output limits.

```yaml
Type: EmailMailboxWriterOptions
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Path
Destination mbox path.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Email.EmailMailbox`

## OUTPUTS

- `OfficeIMO.Email.EmailWriteResult`

## RELATED LINKS

- None
