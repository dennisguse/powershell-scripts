<#
.AUTHOR
Dennis Guse, dennis.guse@alumni.tu-berlin.de

.SYNOPSIS

Executes a SQL query/command, converts (single line, without seperators), and prints the results to STDOUT.
This script is designed for queries with one row and one column.

Requires the Microsoft SQLServer Powershell module: https://docs.microsoft.com/en-us/sql/powershell/sql-server-powershell?view=sql-server-2017
#>


Param(
  [Parameter(Mandatory=$True,Position=1,ValueFromPipeline=$True,ValueFromPipelinebyPropertyName=$True)]
  [string]$serverInstance,
  [Parameter(Mandatory=$True,Position=2,ValueFromPipeline=$True,ValueFromPipelinebyPropertyName=$True)]
  [string]$database,
  [Parameter(Mandatory=$True,Position=3,ValueFromPipeline=$True,ValueFromPipelinebyPropertyName=$True)]
  [string]$sqlQuery
)

Write-Host $MyInvocation.MyCommand.Name

Write-Host $serverInstance
Write-Host $database
Write-Host $sqlQuery

try {
  $resultSet = Invoke-Sqlcmd -ServerInstance ($serverInstance) -Database ($database) -Query ($sqlQuery)
} catch {
  exit -1
}

#Convert resultSet to array[string]
if ($resultSet -is [array]) {
    $resultStringArray =$resultSet.forEach({$_.itemArray })
}
else {
    $resultStringArray = $resultSet.itemArray
}

#Convert table to string (one line)
$resultString = [string]::Concat($resultStringArray)

#Prepare pipeline content
$pipelineOutput = New-Object –typename PSObject
$pipelineOutput | Add-Member -membertype NoteProperty -name SQLResult -value ($resultString)
Write-Output $pipelineOutput

