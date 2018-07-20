<#
.AUTHOR
Dennis Guse, dennis.guse@alumni.tu-berlin.de

.SYNOPSIS

Converts all Office-files (i.e., doc[x], ppt[x], and xls[x]) to PDF.
All files in the working directory (i.e., current directory) are processed.
OUtput files are overwritten.

#>

$Excel = New-Object -ComObject Excel.Application
$Powerpoint = New-Object -ComObject Powerpoint.Application
$Word = New-Object -ComObject Word.Application

#Should warning or error dialogs be shown?
$Powerpoint.DisplayAlerts = [Microsoft.Office.Interop.PowerPoint.PpAlertLevel]::ppAlertsNone

Write-Host $MyInvocation.MyCommand.Name

$inputFilenames = Get-ChildItem | Where-Object{
    $_.Extension -ceq ".xls" -or $_.Extension -ceq ".xlsx" -or $_.Extension -ceq ".ppt" -or $_.Extension -ceq ".pptx" -or $_.Extension -ceq ".doc" -or $_.Extension -ceq ".docx"
} | select -ExpandProperty FullName
$workFolder = Split-Path $MyInvocation.MyCommand.Path

If ($inputFilenames -eq $Null) {
  Write-Warning "No Office-files (i.e., doc[x], ppt[x], and xls[x]) found in folder $workFolder"

  CMD /C PAUSE #Powershell v1.0
  Exit
}

Write-Host "Going to conver t all Office-files from folder $workFolder to PDF`r`n"
$inputFilenames
Write-Host
CMD /C PAUSE #Powershell v1.0
Write-Host


$progress = 0
ForEach ($inputFilename in $inputFilenames) {
  $percentage = $progress++ * 100 / $inputFilenames.Count
  Write-Progress -Activity "progress" -Status $inputFilename -PercentComplete $percentage
  Write-Host $inputFilename

  $outputFilename = (Get-Item $inputFilename).basename + ".pdf"
  $outputPath = Join-Path -Path $workFolder -ChildPath $outputFilename

  If ($inputFilename -like "*.doc" -or $inputFilename -like "*.docx") {
    $inputDocument = $Word.Documents.open($inputFilename, $True, $True)

    Try {
      $inputDocument.SaveAs($outputPath, [Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatPDF)
    } Catch {
      Write-Warning "Saving output to $outputPath failed"
    }
  }


  If ($inputFilename -like "*.ppt" -or $inputFilename -like "*.pptx") {
    $inputDocument = $Powerpoint.Presentations.open($inputFilename, $True, $True, $False)

    Try {
      $inputDocument.SaveAs($outputPath, [Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType]::ppSaveAsPDF)
    } Catch {
      Write-Warning "Saving output to $outputPath failed"
    }
  }

  If ($inputFilename -like "*.xls" -or $inputFilename -like "*.xlsx") {
    $inputDocument = $Excel.Workbooks.open($inputFilename, $True, $True)
    
    Try {
      $inputDocument.ExportAsFixedFormat([Microsoft.Office.Interop.Excel.XlFixedFormatType]::xlTypePDF, $outputPath)
    } Catch {
      Write-Warning "Saving output to $outputPath failed"
    }
  }

  $inputDocument.Close()
}

Write-Progress -Activity "progress" -Status "Completed" -Completed

$Excel.Quit()
$Powerpoint.Quit()
$Word.Quit()

Write-Host "`nDone."
CMD /C PAUSE #Powershell v1.0
