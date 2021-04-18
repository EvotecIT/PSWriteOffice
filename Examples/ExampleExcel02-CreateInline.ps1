Clear-Host
Import-Module .\PSWriteOffice.psd1 -Force

New-OfficeExcel -FilePath "C:\Support\GitHub\PSWriteOffice\Examples\Documents\Excel1.xlsx" {
    # Get-OfficeExcelWorkSheet will try to get workshet with that name, but if not found it will not execute everything inside
    Get-OfficeExcelWorkSheet -Name 'Contact1' {
        # this code will only execute if Contact1 exists
        New-OfficeExcelValue -Row 2 -Column 2 -Value 1
        New-OfficeExcelValue -Row 2 -Column 3 -Value (Get-Date)
        New-OfficeExcelValue -Row 2 -Column 4 -Value 'ok - lets do this'
        New-OfficeExcelValue -Row 2 -Column 5 -Value $true
        New-OfficeExcelValue -Row 2 -Column 6 -Value 'TRUE'

        $Cell = Get-OfficeExcelValue -Row 2 -Column 6
        if ($Cell.DataType -eq 'Boolean') {
            Write-Color -Color Red -Text "Good"
        }
    }
    # New-OfficeExcelWorkSheet will create Contact2 if it doesn't exists, but if it exists, it will use it for it's work
    New-OfficeExcelWorkSheet -Name 'Contact2' {
        New-OfficeExcelValue -Row 2 -Column 2 -Value 1
        New-OfficeExcelValue -Row 2 -Column 3 -Value (Get-Date)
        New-OfficeExcelValue -Row 2 -Column 4 -Value 'ok - lets do this'
        New-OfficeExcelValue -Row 2 -Column 5 -Value $true
        New-OfficeExcelValue -Row 2 -Column 6 -Value 'TRUE'

        $Cell = Get-OfficeExcelValue -Row 2 -Column 6
        if ($Cell.DataType -eq 'Boolean') {
            Write-Color -Color Red -Text "Good Contact 2 exists and we set value"
        }

        $Data1 = Get-Process | Select-Object -First 5

        $Data = @(
            [PSCustomObject] @{
                Test     = 1
                TimeSpan = [timespan]::new(15, 15, 30)
                TimeSpan1 = [timespan]::MinValue
            }
        )

        New-OfficeExcelTable -DataTable $Data1 -Row 10 -Column 15
    }
} -Show -Save