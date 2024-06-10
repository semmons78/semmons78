#
# Script to clear the Teams cache
#

CLEAR

# Load all the cache locations that need to be cleared.
$cacheFolders = @()

$cacheFolders += Join-Path -Path $env:LOCALAPPDATA -ChildPath "packages\msteams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams\*"


# Check if Teams is running
$teamsProcesses = Get-Process -Name "*Teams" -ErrorAction SilentlyContinue
$countTeams = ($teamsProcesses | Measure-Object).Count
Write-Host "Count of Teams processes: $($countTeams)"

if($countTeams -gt 0)
{
    Write-Host "You must close all Teams processes for the cache clear to be performed." -ForegroundColor Red
    $yestNo = Read-Host "Would you like this script to force close all Teams proceses? (y/n)"
    
    if("y" -eq $yestNo -or "yes" -eq $yestNo)
    {
        Write-Host "Closing Teams..."
        Stop-Process -Name "*Teams" -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 10
    }
}

$teamsProcesses = Get-Process -Name "Teams" -ErrorAction SilentlyContinue
$countTeams = ($teamsProcesses | Measure-Object).Count

if($countTeams -gt 0)
{
    Write-Host "Unable to close all Teams processes. Script ending." -ForegroundColor Red
}
else
{
    $fileCount = 0
    $successCount = 0
    $errorCount = 0
    foreach($cache in $cacheFolders)
    {
        if(Test-Path -Path $cache)
        {
            $filesToDelete = Get-ChildItem -Path $cache -Recurse -ErrorAction SilentlyContinue
            
            foreach($f in $filesToDelete)
            {
                if( -not $f.PSIsContainer)
                {
                    $fileCount++
                    try
                    {
                        Write-Host "Deleting file: $($f.FullName)"
                        Remove-Item $($f.FullName) -Force -ErrorAction Stop
                        Write-Host "`tSuccess" -ForegroundColor Green
                        $successCount++
                    }
                    catch
                    {
                        Write-Host "`tERROR: $($_)" -ForegroundColor Red
                        $errorCount++
                    }
                }
            }
        }
    }
    
    Write-Host ""
    Write-Host "Results:"
    Write-Host "Count of files to remove = $($fileCount)"
    Write-Host "Count of success = $($successCount)"
    Write-Host "Count of errors = $($errorCount)"
    Write-Host ""
}

Read-Host "Process complete. Press Enter to exit the script"
