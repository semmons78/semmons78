#
# Script to add Outlook Presence Files
#

CLEAR

# Check if Teams is running
$teamsProcesses = Get-Process -Name "*Teams" -ErrorAction SilentlyContinue
$countTeams = ($teamsProcesses | Measure-Object).Count
Write-Host "Count of Teams processes: $($countTeams)"

if($countTeams -gt 0)
{
    Write-Host "You must close all Teams processes for the Outlook Presence to be fixed." -ForegroundColor Red
    $yestNo = Read-Host "Would you like this script to force close all Teams proceses? (y/n)"
    
    if("y" -eq $yestNo -or "yes" -eq $yestNo)
    {
        Write-Host "Closing Teams..."
        Stop-Process -Name "*Teams" -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 10
    }
}


# Get all Outlook processes
$outlookProcesses = Get-Process Outlook -ErrorAction SilentlyContinue
Write-Host "Outlook processes will now be forced closed to perform Outlook Presence fix." -ForegroundColor Red

# Check if any Outlook processes are found
if ($outlookProcesses) {
    # Close all Outlook processes
    $outlookProcesses | ForEach-Object { $_.CloseMainWindow() | Out-Null }
    Write-Host "Outlook processes have been closed."
} else {
    # No Outlook processes found
    Write-Host "No Outlook processes are running."
}

# Check if TeamsPresenceAddin Folder is Present and add tlb files if needed

$parentPath = $env:LOCALAPPDATA
$childPath ="Microsoft\TeamsPresenceAddin"

$destinationPath = Join-Path -Path $parentPath -ChildPath $childPath


# Check if the folder exists
if (-not (Test-Path -Path $destinationPath)) {
    Write-Host "Creating folder: $($destinationPath)"
    New-Item -ItemType Directory -Path $destinationPath | Out-Null
}

# Copy tlb files to Destination Folder

$source = "\\PCCS01\O365ClientScripts$\Outlook_Presence_Fix\*"
Copy-Item -Path $source -Destination $destinationPath
Write-Host "Folder created. Presence files added" -ForegroundColor Red

Write-Host "Process complete. Script can be closed." -ForegroundColor Red

Read-Host -Prompt "Press enter to close"
