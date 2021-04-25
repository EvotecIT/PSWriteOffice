Clear-Host
Import-Module .\PSWriteOffice.psd1 -Force

$Excel = New-OfficeExcel -FilePath "C:\Support\GitHub\PSWriteOffice\Examples\Documents\Excel2.xlsx" -WhenExists Overwrite

$Data = @(
    [PSCustomObject] @{
        Test     = 1
        TimeSpan = [timespan]::new(15, 15, 30)
        Date     = Get-Date
    }
    [PSCustomObject] @{
        Test      = 2
        TimeSpan  = [timespan]::new(15, 15, 30)
        Date      = Get-Date
        SomeValue = 15
    }
)

$Worksheet = New-OfficeExcelWorksheet -Excel $Excel -Name 'WorkSheet1'
$Table = New-OfficeExcelTable -DataTable $Data -Worksheet $Worksheet -StartRow 1 -StartCell 1 -AllProperties -ReturnObject
New-OfficeExcelTableOptions -Theme 'TableStyleMedium11' -Table $Table -ShowRowStripes -ShowColumnStripes

$Data = @(
    [ordered] @{
        Test     = 1
        TimeSpan = [timespan]::new(15, 15, 30)
        Date     = Get-Date
    }
    [ordered] @{
        Test       = 2
        TimeSpan   = [timespan]::new(15, 15, 30)
        Date       = Get-Date
        SomeValue  = 15
        SomeValue1 = 15
    }
)

$Worksheet = New-OfficeExcelWorksheet -Excel $Excel -Name 'WorkSheet2'
$Table = New-OfficeExcelTable -DataTable $Data -Worksheet $Worksheet -StartRow 1 -StartCell 1 -AllProperties -ReturnObject
New-OfficeExcelTableOptions -Theme 'TableStyleMedium11' -Table $Table -ShowRowStripes -ShowColumnStripes

$Data = Get-Process | Select-Object -First 100

$Worksheet = New-OfficeExcelWorksheet -Excel $Excel -Name 'WorkSheet3'
$Table = New-OfficeExcelTable -DataTable $Data -Worksheet $Worksheet -StartRow 1 -StartCell 1 -AllProperties -ReturnObject
New-OfficeExcelTableOptions -Theme 'TableStyleMedium11' -Table $Table -ShowRowStripes -ShowColumnStripes

$Data = @(
    1
    2
    3
    'test'
)

$Worksheet = New-OfficeExcelWorksheet -Excel $Excel -Name 'WorkSheet4'
$Table = New-OfficeExcelTable -DataTable $Data -Worksheet $Worksheet -StartRow 1 -StartCell 1 -AllProperties -ReturnObject
New-OfficeExcelTableOptions -Theme 'TableStyleMedium11' -Table $Table -ShowRowStripes -ShowColumnStripes

Save-OfficeExcel -Excel $Excel -Show