Param (
    [Parameter(Mandatory)]
    [string]$GlobalUserList,

    [Parameter(Mandatory)]
    [string]$SecurePWPath
)

# Load valid Employee Numbers from CSV
$validEmployeeNumbers = (Import-Csv $GlobalUserList).EmployeeNumber

# Get secure mail password
$securePassword = Get-Content $SecurePWPath | ConvertTo-SecureString -AsPlainText -Force
$EmailCredential = New-Object System.Management.Automation.PSCredential("CorporateEmail@corp.com", $securePassword)

# Email settings
$EmailFrom = "CorporateEmail@corp.com"
$EmailTo = "TechStaff@corp.com"
$SMTPServer = "smtp.office365.com"

# Define rogue log path
$logPath = "$PSScriptRoot\RogueAccounts-$(Get-Date -Format 'yyyy-MM-dd_HHmm').log"

# Search AD for enabled users not in the HR list
$rogueUsers = Get-ADUser -Filter "Enabled -eq $true" -Properties EmployeeNumber | Where-Object {
    $_.EmployeeNumber -and ($_.EmployeeNumber -notin $validEmployeeNumbers)
}

if ($rogueUsers.Count -eq 0) {
    Write-Host "✅ No rogue accounts found."
    return
}

# Disable and log rogue accounts
foreach ($user in $rogueUsers) {
    try {
        Disable-ADAccount -Identity $user.SamAccountName
        $logLine = "[$(Get-Date)] Disabled rogue account: $($user.SamAccountName) ($($user.Name))"
        Add-Content -Path $logPath -Value $logLine
    } catch {
        $errorMsg = "[$(Get-Date)] ERROR disabling $($user.SamAccountName): $_"
        Add-Content -Path $logPath -Value $errorMsg
    }
}

# Compose one summary email
$subject = "⚠ Rogue Accounts Detected: $($rogueUsers.Count) accounts disabled"
$body = Get-Content $logPath -Raw

Send-MailMessage -From $EmailFrom -To $EmailTo -Subject $subject -Body $body -SmtpServer $SMTPServer -Credential $EmailCredential -UseSsl -Port 587

Write-Host "Disabled $($rogueUsers.Count) rogue accounts. Email sent."
