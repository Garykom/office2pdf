$input_file = $args[0]
$output_file = $args[1]

#Write-Output "xls2pdf.ps1 $input_file $output_file"

$ExcelApp = New-Object -ComObject Excel.Application
$ExcelApp.DisplayAlerts = $false
$ExcelApp.Visible = $false

$Workbook = $ExcelApp.workbooks.open($input_file, 3)
$Workbook.Saved = $true

$Workbook.ExportAsFixedFormat([Microsoft.Office.Interop.Excel.xlFixedFormatType]::xlTypePDF, $output_file)

$ExcelApp.Workbooks.Close()
$ExcelApp.Quit()

[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ExcelApp)

$ExcelApp = $null
	
# Make sure references to COM objects are released, otherwise powerpoint might not close
# (calling the methods twice is intentional, see https://msdn.microsoft.com/en-us/library/aa679807(office.11).aspx#officeinteroperabilitych2_part2_gc)
[System.GC]::Collect();
[System.GC]::WaitForPendingFinalizers();
[System.GC]::Collect();
[System.GC]::WaitForPendingFinalizers();
