<#
.AUTHOR
Dennis Guse, dennis.guse@alumni.tu-berlin.de

.SYNOPSIS

Concats the worksheets of all Excel-files (i.e., csv, xls, and xlsx) into one xlsx-file.
All files in the working directory (i.e., current directory) are processed.

.PARAMETER outputFilename The filename of the output file (without extension).
#>
param(
  [string]$outputFilename = "0001-COMBINED"
)

$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $False
#Should confirmation, warning or error dialogs be shown?
$Excel.DisplayAlerts = $True

Write-Host $MyInvocation.MyCommand.Name

$inputFilenames = Get-ChildItem | Where-Object{$_.Extension -ceq ".csv" -or $_.Extension -ceq ".xls" -or $_.Extension -ceq ".xlsx"} | select -ExpandProperty FullName
$workFolder = Split-Path $MyInvocation.MyCommand.Path

$outputPath = Join-Path -Path $workFolder -ChildPath $outputFilename

If ($inputFilenames -eq $Null) {
  Write-Warning "No Excel-files (i.e., csv, xls, and xlsx) found in folder $workFolder"

  CMD /C PAUSE #Powershell v1.0
  Exit
}

Write-Host "Going to concatenate Excel-files (i.e., csv, xls, and xlsx) from folder $workFolder`r`n"
$inputFilenames
Write-Host "`r`nThe result will be stored into $outputFilename`r`n"
CMD /C PAUSE #Powershell v1.0
Write-Host

#Open up a new workbook
$outputWorkbook = $Excel.Workbooks.Add()

$progress = 0
ForEach ($inputFilename in $inputFilenames) {
  $percentage = $progress++ * 100 / $inputFilenames.Count
  Write-Progress -Activity "progress" -Status $inputFilename -PercentComplete $percentage

  $inputWorkbook = $Excel.Workbooks.Open($inputFilename, $True, $True)
  $inputWorkbookBasename = (Get-Item $inputFilename).basename

  ForEach ($inputWorksheet in $inputWorkbook.Worksheets) {
    If ($inputWorkbookBasename -eq $inputWorksheet.Name) {
      $outputWorksheetName = $inputWorkbookBasename
    } else {
      $outputWorksheetName = $inputWorksheet.Name + " - " + $inputWorkbookBasename
    }

    $outputWorksheetNameShort = $outputWorksheetName.Substring(0, [System.Math]::Min(31, $outputWorksheetName.Length))

    If ($inputWorkbookBasename -eq $inputWorksheet.Name) {
      Write-Host $inputWorksheet.Name ">" $outputWorksheetNameShort
    } Else {
      Write-Host $inputWorkbookBasename":" $inputWorksheet.Name ">" $outputWorksheetNameShort
    }

    If ($outputWorksheetNameShort.Length -lt $outputWorksheetName.Length) {
      Write-Warning "Excel only supports Worksheet names up to 31 character: $outputWorksheetNameShort is used"
    }

    $inputWorksheet.Copy([System.Reflection.Missing]::Value, $outputWorkbook.ActiveSheet)

    Try {
      $outputWorkbook.ActiveSheet.Name = $outputWorksheetNameShort
    } Catch {
      Write-Warning "Worksheet name already taken: using default name"
    }
  }

  $inputWorkbook.Close()
}

Write-Progress -Activity "progress" -Status "Saving to $outputPath" -Completed

#Cleanup - remove first (initial and empty) worksheet
$outputWorkbook.Sheets.Item(1).Delete()

Try {
  $outputWorkbook.SaveAs($outputPath)
  Write-Host "Saved to $($outputWorkbook.FullName)"
} Catch {
  Write-Warning "Saving output to $($outputWorkbook.FullName) failed"
}

$Excel.DisplayAlerts = $False #Hide clipboard warning.
$outputWorkbook.close()
$Excel.Quit()

Write-Host "`nDone."
CMD /C PAUSE #Powershell v1.0
