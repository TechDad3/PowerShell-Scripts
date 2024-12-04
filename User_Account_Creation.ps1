# Import Active Directory module
Import-Module ActiveDirectory

# Password generation without System.Web
function New-SecureRandomPassword {
    $lowercase = -join ((97..122) | Get-Random -Count 5 | % {[char]$_})
    $uppercase = -join ((65..90) | Get-Random -Count 3 | % {[char]$_})
    $numbers = -join ((48..57) | Get-Random -Count 4 | % {[char]$_})
    
    $password = -join (($lowercase + $uppercase + $numbers).ToCharArray() | Get-Random -Count 12)
    $securePassword = ConvertTo-SecureString -String $password -AsPlainText -Force
    
    return @{
        SecurePassword = $securePassword
        PlainPassword = $password
    }
}

function New-Username {
    param(
        [string]$FirstName,
        [string]$LastName
    )
    
    $firstName = $FirstName.ToLower() -replace '[^a-zA-Z]', ''
    $lastName = $LastName.ToLower() -replace '[^a-zA-Z]', ''
    $firstInitial = $firstName.Substring(0,1)
    $baseUsername = "$firstInitial$lastName"
    $username = $baseUsername
    
    $counter = 1
    while (Get-ADUser -Filter "SamAccountName -eq '$username'" -ErrorAction SilentlyContinue) {
        $username = "$baseUsername$counter"
        $counter++
    }
    
    return $username
}

function Get-DepartmentOUPath {
    param([string]$Department)
    
    $domainBase = "DC=corp,DC=" ",DC=local"
    
    switch ($Department.ToLower()) {
        "Adult Services" { return "OU=Members,OU=Adult Services,OU=NLG,$domainBase" }
        "Day School" { return "OU=Members,OU=Day School,OU=NLG,$domainBase" }
        default { throw "Invalid department: $Department. Valid departments: Adult Services, Day School" }
    }
}

function New-ADUsersFromCSV {
    param([Parameter(Mandatory=$true)][string]$CSVPath)
    
    $successCount = 0
    $failureCount = 0
    $userReport = @()
    
    try {
        if (-not (Test-Path $CSVPath)) { throw "CSV file not found: $CSVPath" }
    
        $users = Import-Csv -Path $CSVPath
        
        $requiredColumns = @('FirstName', 'LastName', 'Department')
        $csvColumns = ($users | Get-Member -MemberType NoteProperty).Name
        $missingColumns = $requiredColumns | Where-Object { $_ -notin $csvColumns }
        
        if ($missingColumns) {
            throw "Missing columns: $($missingColumns -join ', ')"
        }
        
        foreach ($user in $users) {
            try {
                if (-not ($user.FirstName -and $user.LastName -and $user.Department)) {
                    throw "Missing required fields for user: $($user.FirstName) $($user.LastName)"
                }
                
                $username = New-Username -FirstName $user.FirstName -LastName $user.LastName
                $passwordInfo = New-SecureRandomPassword
                $ouPath = Get-DepartmentOUPath -Department $user.Department
                
                New-ADUser -Name "$($user.FirstName) $($user.LastName)" `
                          -GivenName $user.FirstName `
                          -Surname $user.LastName `
                          -SamAccountName $username `
                          -UserPrincipalName "$username@corp." ".local" `
                          -Department $user.Department `
                          -Path $ouPath `
                          -AccountPassword $passwordInfo.SecurePassword `
                          -Enabled $true `
                          -ChangePasswordAtLogon $true
                
                $userReport += [PSCustomObject]@{
                    FullName = "$($user.FirstName) $($user.LastName)"
                    Username = $username
                    Department = $user.Department
                    OU = $ouPath
                    Password = $passwordInfo.PlainPassword
                    Status = "Success"
                }
                
                Write-Host "Created user: $($user.FirstName) $($user.LastName)" -ForegroundColor Green
                Write-Host "Username: $username" -ForegroundColor Green
                Write-Host "Password: $($passwordInfo.PlainPassword)" -ForegroundColor Green
                Write-Host "Department: $($user.Department)" -ForegroundColor Green
                Write-Host "-------------------"
                
                $successCount++
                
            } catch {
                Write-Host "Failed to create user: $($user.FirstName) $($user.LastName)" -ForegroundColor Red
                Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
                Write-Host "-------------------"
                
                $userReport += [PSCustomObject]@{
                    FullName = "$($user.FirstName) $($user.LastName)"
                    Username = ""
                    Department = $user.Department
                    OU = ""
                    Password = ""
                    Status = "Failed - $($_.Exception.Message)"
                }
                
                $failureCount++
            }
        }
        
        $reportPath = Join-Path (Split-Path $CSVPath) "UserCreationReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        $userReport | Export-Csv -Path $reportPath -NoTypeInformation
        
        Write-Host "`nSummary:" -ForegroundColor Cyan
        Write-Host "Successfully created: $successCount users" -ForegroundColor Green
        Write-Host "Failed to create: $failureCount users" -ForegroundColor Red
        Write-Host "Report exported to: $reportPath" -ForegroundColor Yellow
        
    } catch {
        Write-Host "Error processing CSV file: $($_.Exception.Message)" -ForegroundColor Red
    }
}