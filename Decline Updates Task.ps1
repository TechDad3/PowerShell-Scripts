# Create a new scheduled task to run the Decline Superseded Updates script at 1pm today

# Define the script path
$scriptPath = "C:\PowerShell Scripts\Decline Superseded Updates.ps1"

# Create the task
$taskName = "Decline Superseded Updates"
$taskDescription = "Automatically declines superseded Windows updates"
$taskTrigger = New-ScheduledTaskTrigger -Daily -At 13:38
$taskAction = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-File `"$scriptPath`""
$taskSettings = New-ScheduledTaskSettingsSet

Register-ScheduledTask -TaskName $taskName -Description $taskDescription -Trigger $taskTrigger -Action $taskAction -Settings $taskSettings -Force