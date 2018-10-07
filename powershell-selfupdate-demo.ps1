<#
.AUTHOR
Dennis Guse, dennis.guse@alumni.tu-berlin.de

.SYNOPSIS

Demo of auto-update functionality.
If the file at $updatePath is newer, then the currently executed file is overwritten.

NOTE: by default this script check if itself is newer - execute with a different filename.

#>

Function Update-Check($updatePath) {
    $timestamp = (Get-Item $PSCommandPath).LastWriteTime
    $timestampUpdate = (Get-Item $updatePath).LastWriteTime

    If ($timestamp -lt $timestampUpdate) {
        Write-Host "Updating script $(Get-Item $PSCommandPath) to version from $timestampUpdate"
        Pause

        Copy-Item -Path $updatePath -Destination (Get-Item $PSCommandPath)

        Write-Host "Update successfully installed."
        Write-Host "ATTENTION: This program will terminate now. Please restart the program."
        Pause
        Exit
    }
}

Update-Check "powershell-selfupdate-demo.ps1"
