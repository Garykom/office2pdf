$input_file = $args[0]
$output_file = $args[1]

#Write-Output "ppt2pdf.ps1 $input_file $output_file"

$PowerPointApp = New-Object -ComObject PowerPoint.Application

$Presentation = $PowerPointApp.Presentations.Open($input_file)
$Presentation.Saved = $true

$Presentation.SaveAs($output_file, [Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType]::ppSaveAsPDF)

$Presentation.Close()
$Presentation = $null
	
if ($PowerPointApp.Windows.Count -eq 0) {
    $PowerPointApp.Quit()
}
	
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($PowerPointApp)

$PowerPointApp = $null
	
# Make sure references to COM objects are released, otherwise powerpoint might not close
# (calling the methods twice is intentional, see https://msdn.microsoft.com/en-us/library/aa679807(office.11).aspx#officeinteroperabilitych2_part2_gc)
[System.GC]::Collect();
[System.GC]::WaitForPendingFinalizers();
[System.GC]::Collect();
[System.GC]::WaitForPendingFinalizers();
