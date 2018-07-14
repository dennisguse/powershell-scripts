<#
.AUTHOR
Dennis Guse, dennis.guse@alumni.tu-berlin.de

.SYNOPSIS

Converts all Powerpoint-files (i.e., ppt and pptx) to PDF.
All files in the working directory (i.e., current directory) are processed.

#>

$Powerpoint = New-Object -ComObject powerpoint.Application
 
#$Powerpoint.Visible = [Microsoft.Office.Core.MsoTriState]::msoFalse
#Should warning or error dialogs be shown?
$Powerpoint.DisplayAlerts = [Microsoft.Office.Interop.PowerPoint.PpAlertLevel]::ppAlertsNone

Write-Host $MyInvocation.MyCommand.Name

$inputFilenames = Get-ChildItem | Where-Object{$_.Extension -ceq ".ppt" -or $_.Extension -ceq ".pptx"} | select -ExpandProperty FullName
$workFolder = Split-Path $MyInvocation.MyCommand.Path

If ($inputFilenames -eq $Null) {
  Write-Warning "No Powerpoint-files (i.e., ppt and pptx) found in folder $workFolder"

  CMD /C PAUSE #Powershell v1.0
  Exit
}

Write-Host "Going to convert all Powerpoint-files from folder $workFolder to PDF`r`n"
$inputFilenames
Write-Host
CMD /C PAUSE #Powershell v1.0
Write-Host


$progress = 0
ForEach ($inputFilename in $inputFilenames) {
  $percentage = $progress++ * 100 / $inputFilenames.Count
  Write-Progress -Activity "progress" -Status $inputFilename -PercentComplete $percentage

  $inputPresentation = $Powerpoint.Presentations.open($inputFilename, $True, $True)
  $inputPresentationBasename = (Get-Item $inputFilename).basename

  Try {
    $inputPresentation.SaveAs($inputFilename, [Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType]::ppSaveAsPDF)
    $inputPresentation.Close()
  } Catch {
    Write-Warning "Saving output to $outputPath failed"
  }

  $inputPresentation.Close()
}

Write-Progress -Activity "progress" -Status "Completed" -Completed

$Powerpoint.Quit()

Write-Host "`nDone."
CMD /C PAUSE #Powershell v1.0
