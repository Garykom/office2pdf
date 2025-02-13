$input_file = $args[0]
$output_file = $args[1]

#Write-Output "doc2pdf.ps1 $input_file $output_file"

$WordApp = New-Object -ComObject Word.Application
$WordApp.Visible = $false

$Document = $WordApp.Documents.Open($input_file)
$Document.SaveAs([ref] $output_file, [ref] 17)

$Document.Close()
$WordApp.Quit()

[System.Runtime.InteropServices.Marshal]::ReleaseComObject($WordApp)

$WordApp = $null
	
# Make sure references to COM objects are released, otherwise powerpoint might not close
# (calling the methods twice is intentional, see https://msdn.microsoft.com/en-us/library/aa679807(office.11).aspx#officeinteroperabilitych2_part2_gc)
[System.GC]::Collect();
[System.GC]::WaitForPendingFinalizers();
[System.GC]::Collect();
[System.GC]::WaitForPendingFinalizers();
