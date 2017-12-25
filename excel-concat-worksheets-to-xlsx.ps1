<#
.AUTHOR
Dennis Guse, dennis.guse@alumni.tu-berlin.de

.SYNOPSIS

Concats the worksheets of all Excel-files (i.e., csv, xls, and xlsx) into one xlsx-file.
Worksheets are concatenated using their _index_ in the original file.
All files in the working directory (i.e., current directory) are processed.

.NOTES
The name of worksheet is the same as the _last_ worksheet (with this index) to be concatenated.

.PARAMETER outputFilename The filename of the output file (without extension).
#>
param(
  [string]$outputFilename = "0001-COMBINED"
)

$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $False
#Should warning or error dialogs be shown?
$Excel.DisplayAlerts = $True

Write-Host $MyInvocation.MyCommand.Name

$inputFilenames = Get-ChildItem | Where-Object{$_.Extension -ceq ".csv" -or $_.Extension -ceq ".xls" -or $_.Extension -ceq ".xlsx"} | select -ExpandProperty FullName
$workFolder = Split-Path $MyInvocation.MyCommand.Path

$outputPath = Join-Path -Path $workFolder -ChildPath $outputFilename

If ($inputFilenames -eq $Null) {
  Write-Warning "No Excel-files (i.e., csv, xls, and xlsx) found in folder $workFolder"

  Pause
  Exit
}

Write-Host "Going to concatenate Excel-files (i.e., csv, xls, and xlsx) from folder $workFolder`r`n"
$inputFilenames
Write-Host "`r`nThe result will be saved as $outputFilename`r`n"
Pause
Write-Host

#Open up a new workbook
$outputWorkbook = $Excel.Workbooks.Add()

$progress = 0
ForEach ($inputFilename in $inputFilenames) {
  $percentage = $progress++ * 100 / $inputFilenames.Count
  Write-Progress -Activity "progress" -Status $inputFilename -PercentComplete $percentage

  $inputWorkbook = $Excel.Workbooks.Open($inputFilename, $True, $True)
  $inputWorkbookBasename = (Get-Item $inputFilename).basename

  $inputWorksheetIndex = 0
  ForEach ($inputWorksheet in $inputWorkbook.Worksheets) {
    $inputWorksheetIndex++
    $outputWorksheetNameShort = $inputWorksheet.Name.Substring(0, [System.Math]::Min(31, $inputWorksheet.Name.Length))

    If ($inputWorkbookBasename -eq $inputWorksheet.Name) {
      Write-Host $inputWorksheet.Name "> Sheet" $inputWorksheetIndex
    } Else {
      Write-Host $inputWorkbookBasename":" $inputWorksheet.Name "> Sheet" $inputWorksheetIndex
    }

    If ($outputWorksheetNameShort.Length -lt $outputWorksheetName.Length) {
      Write-Warning "Excel only supports Worksheet names with up to 31 character: $outputWorksheetNameShort is used"
    }

    [void]$inputWorksheet.UsedRange.Copy()

    $outputWorksheets = $outputWorkbook.Sheets
    For ($i = $outputWorkbook.Sheets.Count; $i -lt $inputWorksheetIndex; $i++) {
      [void]$outputWorkbook.Sheets.Add([System.Reflection.Missing]::Value, $outputWorkbook.Sheets($outputWorkbook.Sheets.Count))
    }


    $outputWorksheet = $outputWorkbook.Sheets($inputWorksheetIndex)
    $lastRow = "A$($outputWorksheet.UsedRange.Rows.Count + 1)"
    $range = $outputWorksheet.Range($lastRow)
    $outputWorksheet.Paste($range)

    Try {
      $outputWorksheet.Name = $outputWorksheetNameShort
    } Catch {
      Write-Warning "Worksheet name already taken: using default name"
    }
  }

  $inputWorkbook.Close()
}

Write-Progress -Activity "progress" -Status "Saving to $outputPath" -Completed

Try {
  $outputWorkbook.SaveAs($outputPath)
  Write-Host "Saved to $($outputWorkbook.FullName)"
} Catch {
  Write-Warning "Saving output to $($outputWorkbook.FullName) failed"
}

$Excel.DisplayAlerts = $False #Hide clipboard warning.
$outputWorkbook.Close()
$Excel.Quit()

Write-Host "`nDone"
Pause
