# Import the Update Services module
Import-Module UpdateServices

# Connect to the WSUS server
$wsus = Get-WsusServer -Name "server name" -Port 8530

# Get all updates
$updates = $wsus.GetUpdates()

# Get the date one week ago
$oneWeekAgo = (Get-Date).AddDays(-7)

# Get current date
$currentDate = Get-Date

# Filter for superseded updates that are up to one week old
$supersededUpdates = $updates | Where-Object { 
    $_.IsSuperseded -and 
    $_.ArrivalDate -ge $oneWeekAgo -and 
    $_.ArrivalDate -le $currentDate
}

# Add logging
$logPath = "C:\PowerShell Scripts\PowerShell Logs\DeclineUpdates_Log.txt"
$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
Add-Content -Path $logPath -Value "=== Script executed at $timestamp ==="

# Decline each superseded update
$declinedCount = 0
foreach ($update in $supersededUpdates) {
    try {
        $update.Decline()
        $message = "Declined update: $($update.Title) (Arrival Date: $($update.ArrivalDate))"
        Write-Host $message
        Add-Content -Path $logPath -Value $message
        $declinedCount++
    }
    catch {
        $errorMessage = "Error declining update $($update.Title): $($_.Exception.Message)"
        Write-Host $errorMessage -ForegroundColor Red
        Add-Content -Path $logPath -Value $errorMessage
    }
}

# Log summary
$summary = "Total updates declined: $declinedCount"
Write-Host $summary
Add-Content -Path $logPath -Value $summary
Add-Content -Path $logPath -Value "=== Script completed at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ===`n"