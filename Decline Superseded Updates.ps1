# Import the Update Services module
>> Import-Module UpdateServices
>>
>> # Connect to the WSUS server
>> $wsus = Get-WsusServer -Name "server name" -Port 8530
>>
>> # Get all updates
>> $updates = $wsus.GetUpdates()
>>
>> # Get the date one week ago
>> $oneWeekAgo = (Get-Date).AddDays(-7)
>>
>> # Filter for superseded updates within the last week
>> $supersededUpdates = $updates | Where-Object { $_.IsSuperseded -eq $true -and $_.ArrivalDate -ge $oneWeekAgo }
>>
>> # Decline each superseded update
>> foreach ($update in $supersededUpdates) {
>>     $update.Decline()
>>     Write-Host "Declined update: $($update.Title)"
>> }
>>
>> Write-Host "All superseded updates within the last week have been declined."