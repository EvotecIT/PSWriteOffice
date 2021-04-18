Clear-Host
Import-Module .\PSWriteOffice.psd1 -Force

New-OfficeExcel -FilePath "C:\Support\GitHub\PSWriteOffice\Examples\Documents\Excel5.xlsx" {
    New-OfficeExcelWorkSheet -Name 'Contact54455' {
        $Data1 = Get-Process | Select-Object -First 5
        New-OfficeExcelTable -DataTable $Data1 -Row 1 -Column 1 -DisableAutoFilter -Theme TableStyleMedium28  #-EmphasizeFirstColumn
    }
} -Show -Save