<#
.AUTHOR
Dennis Guse, dennis.guse@alumni.tu-berlin.de

.SYNOPSIS

Splits Excel-files (i.e., csv, xls, and xlsx) into seperate CSV-files (one CSV-file per worksheet - RFC4180).
All files in the working directory (i.e., current directory) are processed.
#>

$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $False
#Should warning or error dialogs be shown?
$Excel.DisplayAlerts = $False

Write-Host $MyInvocation.MyCommand.Name

$inputFilenames = Get-ChildItem | Where-Object{$_Extension -ceq ".csv" -or $_.Extension -ceq ".xls" -or $_.Extension -ceq ".xlsx"} | select -ExpandProperty FullName
$workFolder = Split-Path $MyInvocation.MyCommand.Path

If ($inputFilenames -eq $Null) {
  Write-Warning "No Excel-files (i.e., xls and xlsx) found in folder $workFolder"

  CMD /C PAUSE #Powershell v1.0
  Exit
}

Write-Host "Going to split all Excel-files (i.e., xls and xlsx) from folder $workFolder`r`n"
$inputFilenames
Write-Host
CMD /C PAUSE #Powershell v1.0
Write-Host

$progress = 0
ForEach ($inputFilename in $inputFilenames) {
  $percentage = $progress++ * 100 / $inputFilenames.Count
  Write-Progress -Activity "progress" -Status $inputFilename -PercentComplete $percentage

  $inputWorkbook = $Excel.Workbooks.Open($inputFilename, $True, $True)
  $inputWorkbookBasename = (Get-Item $inputFilename).basename

  ForEach ($inputWorksheet in $inputWorkbook.Worksheets) {
    $outputPath = Join-Path -Path $workFolder -ChildPath $($inputWorkbookBasename + "_" + $inputWorksheet.Name)

    Write-Host $inputFilename":" $inputWorksheet.Name ">" $outputPath

    Try {
      $inputWorksheet.SaveAs($outputPath, 6)
    } Catch {
      Write-Warning "Saving output to $outputPath failed"
    }
  }

  $inputWorkbook.Close()
}

Write-Progress -Activity "progress" -Status "Completed" -Completed

$Excel.Quit()

Write-Host "`nDone."
CMD /C PAUSE #Powershell v1.0
