<#
.AUTHOR
Dennis Guse, dennis.guse@alumni.tu-berlin.de

.SYNOPSIS

Extracts the selected column (i.e., defined by a search string) from all Excel-files (i.e., csv, xls, and xlsx) into one seperate Excel-files.
Only the selected column is extracted AND only rows after the search string.
All files in the working directory (i.e., current directory) are processed.

.PARAMETER searchString The string to be searched for.
#>
param(
  [Parameter(Mandatory=$True,Position=1,ValueFromPipeline=$True,ValueFromPipelinebyPropertyName=$True)]
  [string]$searchString
)

$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $False
#Should warning or error dialogs be shown?
$Excel.DisplayAlerts = $False

Write-Host $MyInvocation.MyCommand.Name

$inputFilenames = Get-ChildItem | Where-Object{$_.Extension -ceq ".csv" -or $_.Extension -ceq ".xls" -or $_.Extension -ceq ".xlsx"} | select -ExpandProperty FullName
$workFolder = Split-Path $MyInvocation.MyCommand.Path

If ($inputFilenames -eq $Null) {
  Write-Warning "No Excel-files (i.e., csv, xls, and xlsx) found in folder $workFolder"

  CMD /C PAUSE #Powershell v1.0
  Exit
}

Write-Host "Going to extract one column of the Excel-files from folder $workFolder`r`n"
$inputFilenames
Write-Host "Using search string $searchString"
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
    If ($inputWorkbookBasename -eq $inputWorksheet.Name) {
      $outputWorksheetName = $inputWorkbookBasename + " - extract"
    } else {
      $outputWorksheetName = $inputWorksheet.Name + " - " + $inputWorkbookBasename
    }

    Write-Host $inputWorksheet.Name ">" $outputWorksheetName

    #Select column
    $column = $inputWorksheet.Range("A1:Z4").Find($searchString)

    if ($column -eq $Null) {
      Write-Warning "$($inputFilename): $($inputWorksheet.Name) does not contain $searchString"
      continue
    }

    #Open up a new workbook
    $outputWorkbook = $Excel.Workbooks.Add()

    #Copy&paste
    [void]$column.EntireColumn.Copy()
    $outputWorkbook.ActiveSheet.Paste($outputWorkbook.ActiveSheet.Range("A1"))

    #Delete (header) rows
    [void]$outputWorkbook.ActiveSheet.Range("A1:A$($column.row)").EntireRow.Delete()

    $outputPath = Join-Path -Path $workFolder -ChildPath $outputWorksheetName

    Try {
      $outputWorkbook.SaveAs($outputPath)
      $outputWorkbook.Close()
    } Catch {
      Write-Warning "Saving output to $outputPath failed"
    }
  }

  $inputWorkbook.Close()
}

Write-Progress -Activity "progress" -Status "Completed" -Completed

$Excel.DisplayAlerts = $False #Hide clipboard warning.
$Excel.Quit()

Write-Host "`nDone."
CMD /C PAUSE #Powershell v1.0
