Write-Output "Convert running...";

$script_dir = $args[0]
$input_dir = $args[1]
$output_dir = $args[2]

if ( $null -eq $script_dir) {
    $script_dir = $MyInvocation.MyCommand.Path | split-path -parent
}
if ( $null -eq $input_dir) {
    $input_dir = "$script_dir\input"
}
if ( $null -eq $output_dir) {
    $output_dir = "$script_dir\output"
}

Write-Output "script_dir = $script_dir"
Write-Output "input_dir = $input_dir"
Write-Output "output_dir = $output_dir"

$joblist = New-Object System.Collections.ArrayList

$ping2pdf = {       
    param($script_dir, $input_file, $output_file)
}

$doc2pdf = {       
    param($script_dir, $input_file, $output_file)
    & "$script_dir\doc2pdf.ps1" $input_file $output_file 
    Remove-Item $input_file   
}

$xls2pdf = {       
    param($script_dir, $input_file, $output_file)
    & "$script_dir\xls2pdf.ps1" $input_file $output_file 
    Remove-Item $input_file
}

$ppt2pdf = {       
    param($script_dir, $input_file, $output_file)
    & "$script_dir\ppt2pdf.ps1" $input_file $output_file 
    Remove-Item $input_file
}

$string = ".ping .pdf"
$ping_extensions = $string.Split(" ")

$string = ".doc .docm .docx .dot .dotm .dotx .htm .html .mht .mhtml .odt .rtf .txt .wps .xml .xps"
$word_extensions = $string.Split(" ")

$string = ".csv .dbf .ods .xls .xlsb .xlsm .xlsx .xlt .xltm .xltx .xlw"
$excel_extensions = $string.Split(" ")

$string = ".odp .ppt .pptx"
$powerpoint_extensions = $string.Split(" ")

While ($true) {

    Get-ChildItem -Path $input_dir | ForEach-Object {
    
        $input_file = $_.FullName
        $copy_file = "$($output_dir)\$($_.Name)"
        $output_file = "$($output_dir)\$($_.BaseName).pdf"
        $input_file_extension = $_.Extension

        Write-Output "Copy file: $input_file to: $copy_file"
        Move-Item $input_file -Destination $copy_file

        $command = $null
        if ($ping_extensions.Contains($input_file_extension)) {
            $command = $ping2pdf
        }
        elseif ($word_extensions.Contains($input_file_extension)) {
            $command = $doc2pdf
        }
        elseif ($excel_extensions.Contains($input_file_extension)) {
            $command = $xls2pdf
        }
        elseif ($powerpoint_extensions.Contains($input_file_extension)) {
            $command = $ppt2pdf
        }

        if ($null -eq $command) {
            Write-Output "Unknown Extension: $input_file_ext"
        }
        else {
            Write-Output "Convert $input_file_ext file: $copy_file to: $output_file"
            $jb = Start-Job $command -Args $script_dir, $copy_file, $output_file
            $joblist.Add($jb)
        }
    }

    foreach ($jb in $joblist) {
        $rez = $jb | Wait-Job | Receive-Job
        Write-Output "rez = $rez"
    }

    Write-Output "Sleep 5 sec"
    Start-Sleep -Seconds 5
}
