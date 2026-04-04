# PSWriteOffice Issue Close Templates

These are release-safe response drafts for use only after the relevant version is published.

## Rules

- Confirm the release version first.
- Link to the released module or release notes.
- Do not say an issue is fixed if the change only exists in source or an unreleased branch.
- Keep replies short and concrete.

## Template: Docs / Example

`This is included in the released PSWriteOffice <version>. We added an example covering this scenario here: <link>. Closing this out, but if you still hit a gap on the released build, please reopen with a minimal sample.`

## Template: New Cmdlet / Wrapper

`This is available in the released PSWriteOffice <version>. The supported command path is <cmdlet/example>. Closing this issue based on the released implementation.`

## Template: OfficeIMO-dependent fix

`This is included in the released PSWriteOffice <version>, which consumes the OfficeIMO build containing the related fix. The released path is <cmdlet/example>. Closing this out, but please reopen if you can still reproduce it on the published version.`

## Issue-specific drafts

### `#1` Search and Replace Text

`This is available in the released PSWriteOffice <version> through Update-OfficeWordText (alias Replace-OfficeWordText). Closing this issue based on the released implementation.`

### `#3` Add Chart/Picture into Word Table Cell

`This is supported in the released PSWriteOffice <version>. You can author pictures, lists, and nested tables directly inside table cells, and charts can be anchored to a paragraph created inside a table cell. Closing this out based on the released implementation and examples.`

### `#4` Add table to table cell

`This is supported in the released PSWriteOffice <version>. Nested Word tables can now be created inside table cells from the PowerShell DSL.`

### `#5` Add Line Breaks

`This is covered in the released PSWriteOffice <version> examples. We added a line-break example showing both paragraph breaks and same-paragraph breaks.`

### `#7` No Charts :(

`This is supported in the released PSWriteOffice <version> through Add-OfficeWordChart / WordChart. Closing this issue based on the released chart cmdlet surface.`

### `#8` -Transpose parameter

`This is available in the released PSWriteOffice <version>. Add-OfficeWordTable now supports transpose output.`

### `#10` Set font name

`This behavior is addressed in the released build we ship with PSWriteOffice <version>. Closing this out based on the released font handling behavior.`

### `#12` Underline spaces/tabs

`This is addressed in the released PSWriteOffice <version>, which consumes the related OfficeIMO fix for underline/tab round-tripping. Closing this issue based on the released package combination.`

### `#13` Header / Footer Support

`This is supported in the released PSWriteOffice <version>. Header and footer cmdlets are now part of the module.`

### `#14` Bulleted list within table cell

`This is supported in the released PSWriteOffice <version>. Lists can now be authored inside Word table cells from the DSL.`

### `#15` Null array error when exporting to word

`We added regression coverage around this path and could not reproduce document corruption on the released PSWriteOffice <version>. Closing this based on the released behavior, but please reopen with a minimal reproducer if you still see it on the published build.`

### `#18` Close-OfficeWord without object

`This is improved in the released PSWriteOffice <version>. Close-OfficeWord now has better tracked cleanup ergonomics for the current document flow.`

### `#19` PieChart method

`This is covered in the released PSWriteOffice <version> by the current Add-OfficeWordChart / WordChart path and updated examples for pie charts.`

### `#20` Add some extra columns!?

`This is covered in the released PSWriteOffice <version> examples by projecting object properties and calculated columns before calling Add-OfficeWordTable.`

### `#21` License

`This is addressed in the released PSWriteOffice <version>. The repository now includes a LICENSE file and the module metadata points to it correctly.`

### `#26` TableLayout AutoFit behavior

`This is available in the released PSWriteOffice <version>. Add-OfficeWordTable now exposes the expanded layout options for Word table autofit behavior.`
