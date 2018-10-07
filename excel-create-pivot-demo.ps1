<#
.AUTHOR
Dennis Guse, dennis.guse@alumni.tu-berlin.de

.SYNOPSIS

Adds a Pivot table into an existing Excel table.

NOTE: This is only a demo that needs to be adjusted for the specific use case (e.g., names of data columns).

API documentation is available here: https://msdn.microsoft.com/en-us/vba/excel-vba/articles/object-model-excel-vba-reference
#>
$Excel = New-Object -ComObject excel.application

#Excel constants

#See https://msdn.microsoft.com/en-us/vba/excel-vba/articles/xlsortorder-enumeration-excel
$xlDescending = [Microsoft.Office.Interop.Excel.XLSortOrder]::xlDescending

#See https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel.xlpivottablesourcetype
$xlDatabase = [Microsoft.Office.Interop.Excel.XlPivotTableSourceType]::xlDatabase


#see https://msdn.microsoft.com/en-us/vba/excel-vba/articles/pivotfield-orientation-property-excel
$xlHidden = [Microsoft.Office.Interop.Excel.XlPivotFieldOrientation]::xlHidden
$xlRowField = [Microsoft.Office.Interop.Excel.XlPivotFieldOrientation]::xlRowField
$xlColumnField = [Microsoft.Office.Interop.Excel.XlPivotFieldOrientation]::xlColumnField
$xlPageField = [Microsoft.Office.Interop.Excel.XlPivotFieldOrientation]::xlPageField
$xlDataField = [Microsoft.Office.Interop.Excel.XlPivotFieldOrientation]::xlDataField

#See https://msdn.microsoft.com/en-us/vba/excel-vba/articles/xlconsolidationfunction-enumeration-excel
$xlSum = [Microsoft.Office.Interop.Excel.XlConsolidationFunction]::xlSum
$xlAverage = [Microsoft.Office.Interop.Excel.XlConsolidationFunction]::xlAverage
$xlCount = [Microsoft.Office.Interop.Excel.XlConsolidationFunction]::xlCount
$xlRight = [Microsoft.Office.Interop.Excel.XlConsolidationFunction]::xlSum

#see https://msdn.microsoft.com/en-us/vba/excel-vba/articles/xlpivotfieldcalculation-enumeration-excel
$xlPercentOfColumn = [Microsoft.Office.Interop.Excel.XlPivotFieldCalculation]::xlPercentOfColumn

#See https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel.xlptselectionmode
$xlBlanks = [Microsoft.Office.Interop.Excel.XlPTSelectionMode]::xlBanks
$xlButton = [Microsoft.Office.Interop.Excel.XlPTSelectionMode]::xlButton
$xlDataAndLabel	= [Microsoft.Office.Interop.Excel.XlPTSelectionMode]::xlDataAndLabel
$xlDataOnly	= [Microsoft.Office.Interop.Excel.XlPTSelectionMode]::xlDataOnly
$xlFirstRow	= [Microsoft.Office.Interop.Excel.XlPTSelectionMode]::xlFirstRow
$xlLabelOnly = [Microsoft.Office.Interop.Excel.XlPTSelectionMode]::xlLabelOnly
$xlOrigin = [Microsoft.Office.Interop.Excel.XlPTSelectionMode]::xlOrigin

#Show Excel window
$Excel.Visible = $True

###
### The actual script
###
$path = Resolve-Path ".\testdata\excel-create-pivot-demo\testdata-1.xlsx"
$workbook = $Excel.Workbooks.Open($path, $TRUE, $TRUE)

#Which worksheet contains the data? (Assumption: first)
$sheetData = $workbook.ActiveSheet

#Add pivot table - uses the table that is _activated_ at the moment PivotTableWizard is called.
$pivotTable = $sheetData.PivotTableWizard()

$sheetPivot = $workbook.ActiveSheet

$pivotTable.NullString = "0"
$pivotTable.DisplayNullString = $TRUE

#Columns, rows, and data
$pivotTable.PivotFields("ONE").Orientation = [int]$xlRowField
$pivotTable.PivotFields("TWO").Orientation = [int]$xlColumnField

#Data
$pivotDataSum = $pivotTable.PivotFields("Value")
$pivotDataSum.Orientation = [int]$xlDataField
$pivotDataSum.NumberFormat = "#.##0,00 C"

$pivotDataCount = $pivotTable.PivotFields("Value")
$pivotDataCount.Orientation = [int]$xlDataField
$pivotDataCount.Function = [int]$xlCount

$pivotDataPercent = $pivotTable.PivotFields("Value")
$pivotDataPercent.Orientation = [int]$xlDataField
$pivotDataPercent.Calculation = [int]$xlPercentOfColumn
$pivotDataPercent.NumberFormat = "0,00%"

#Set caption for variables
$pivotDataSum.Caption = "Value (sum)"
$pivotDataCount.Caption = "Value (count)"
$pivotDataPercent.Caption = "EUR (percent)"

#Move Data field to column
$pivotTable.DataPivotField.Orientation = [int]$xlColumnField

#Hide grand total
$pivotTable.RowGrand = $FALSE

#Hide a column (i.e., value)
$pivotTable.PivotFields("ONE").PivotItems("2").visible = $FALSE

#Group a factor (here ONE)
$one_zero = $pivotTable.PivotFields("ONE").PivotItems("0").LabelRange
$one_one = $pivotTable.PivotFields("ONE").PivotItems("1").LabelRange

$b = $sheetPivot.Range($one_zero, $one_one).Group()

##Dirty hack to get the row group (might be locale dependent)
$rowGroup = $pivotTable.PivotFields("ONE2")
$rowGroup.ShowDetail = $FALSE

##Set the group label manually.
$rowGroup.VisibleItems() | Where-Object {$_.ChildItems().Count -gt 0}
FOREACH($item in $rowGroup.VisibleItems() | Where-Object {$_.ChildItems().Count -gt 0}) {
    $currentChildItems = $item.ChildItems()

    $caption = $currentChildItems[1].Value
    FOR($i = 2; $i -le $currentChildItems.Count; $i++) {
        $caption = $caption + " & " + $currentChildItems[$i].Value
    }
    $item.Caption = $caption
}
