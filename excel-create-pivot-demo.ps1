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
$xlPivotTableVersion12  = 3
$xlPivotTableVersion10  = 1

$xlDescending= 2
$xlDatabase                = 1
$xlHidden                  = 0
$xlRowField                = 1
$xlColumnField             = 2
$xlFilterField             = 3
$xlDataField               = 4    
$xlDirection               = [Microsoft.Office.Interop.Excel.XLDirection]

$xlSum                     = -4157
$xlAverage                 = -4106
$xlCount                   = -4112
$xlRight                   = -4152

$xlPercentOfColumn         = 7

#XlPTSelectionMode
$xlBlanks	= 4
$xlButton	= 15
$xlDataAndLabel	= 0 
$xlDataOnly	= 2
$xlFirstRow	= 256
$xlLabelOnly	= 1
$xlOrigin	= 3

#Show Excel window
$Excel.Visible = $True

###
### The actual script
###
$workbook = $Excel.Workbooks.Open("C:\Users\Guse\Desktop\powershell-scripts.new\testdata\excel-create-pivot-demo\testdata-1.xlsx", $True, $True)

#Which worksheet contains the data? (Assumption: first)
$sheetData = $workbook.ActiveSheet

#Add pivot table
$pivotTable = $sheetData.PivotTableWizard()
#
$sheetPivot = $workbook.ActiveSheet

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
$ONE_ZERO = $pivotTable.PivotFields("ONE").PivotItems("0").LabelRange
$ONE_ONE = $pivotTable.PivotFields("ONE").PivotItems("1").LabelRange

$b=$sheetPivot.Range($ONE_ZERO, $ONE_ONE).Group()

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