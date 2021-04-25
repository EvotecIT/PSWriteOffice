Clear-Host
Import-Module .\PSWriteOffice.psd1 -Force

$Data1 = Get-Process | Select-Object -First 15
$Objects = @(
    [PSCustomObject] @{ Test = 1; DateTime = (Get-Date); TimeSpan = (New-TimeSpan -Minutes 10); TestString = 'string' }
    [PSCustomObject] @{ Test = 1; }
    [PSCustomObject] @{ Test = 3; DateTime = (Get-Date).AddDays(1); TimeSpan = (New-TimeSpan -Minutes 10); TestString = 'string' }
)

New-OfficeExcel -FilePath "C:\Support\GitHub\PSWriteOffice\Examples\Documents\Excel.xlsx" {
    New-OfficeExcelWorkSheet -Name 'Contact1' {
        New-OfficeExcelTable -DataTable $Data1 -Row 1 -Column 1 -Theme TableStyleMedium28
    }
    New-OfficeExcelWorkSheet -Name 'Contact2' {
        New-OfficeExcelTable -DataTable $Objects -Row 1 -Column 1 -Theme TableStyleMedium28
    }
    New-OfficeExcelWorkSheet -Name 'Contact3' {
        New-OfficeExcelValue -Row 1 -Column 1 -Value 'Test1'
        New-OfficeExcelValue -Row 1 -Column 2 -Value 'Test2'
        New-OfficeExcelValue -Row 1 -Column 3 -Value 'Test3'
        Set-OfficeExcelValueStyle -Row 1 -Column 3 -FontSize 30 -FontFamilyNumbering Decorative -VerticalAlignment Superscript
        New-OfficeExcelValue -Row 1 -Column 4 -Value 'Test4'
        Set-OfficeExcelValueStyle -Row 1 -Column 4 -Italic $true -Bold $true -BackGroundColor BattleshipGrey
        New-OfficeExcelValue -Row 2 -Column 1 -Value 'Test5'
        New-OfficeExcelValue -Row 2 -Column 2 -Value 'Test6'
        New-OfficeExcelValue -Row 2 -Column 3 -Value 20000
        Set-OfficeExcelValueStyle -Row 2 -Column 3 -FormatID 15 -FontSize 15 -FontColor Blue -Underline DoubleAccounting
        New-OfficeExcelValue -Row 2 -Column 4 -Value 30000 #-Format "$ #,##0"
        Set-OfficeExcelValueStyle -Row 2 -Column 4 -Format "$ #,##0" -FontColor '#0001F1'
    }

} -Show -Save -WhenExists Overwrite