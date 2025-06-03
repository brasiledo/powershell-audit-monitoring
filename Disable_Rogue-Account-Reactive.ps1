Param (
    [string]$GlobalUserList = "C:\Scripts\EmployeeList.csv",
    [string]$SecurePWPath = "C:\Scripts\mailpass.txt"
)

# Load valid Employee Numbers
$validIDs = (Import-Csv $GlobalUserList).EmployeeNumber

# Get last Event ID 4720 (user creation)
$event = Get-WinEvent -LogName Security -FilterXPath "*[System[(EventID=4720)]]" -MaxEvents 1
if (-not $event) { return }

# Extract account name from event XML
[xml]$eventXml = $event.ToXml()
$newUser = $eventXml.Event.EventData.Data | Where-Object { $_.Name -eq 'TargetUserName' } | Select-Object -ExpandProperty '#text'

# Get AD user object
$user = Get-ADUser -Identity $newUser -Properties EmployeeNumber -ErrorAction SilentlyContinue
if (-not $user) { return }

# Validate against list
if ($user.EmployeeNumber -and ($user.EmployeeNumber -in $validIDs)) {
    Write-Host "✅ Valid account created: $newUser"
    return
}

# Disable rogue account
Disable-ADAccount -Identity $user.SamAccountName

# Prepare email
$securePassword = Get-Content $SecurePWPath | ConvertTo-SecureString -AsPlainText -Force
$EmailCredential = New-Object System.Management.Automation.PSCredential("CorporateEmail@corp.com", $securePassword)

$body = "Rogue account disabled:`n`nUser: $($user.SamAccountName)`nCreated: $($event.TimeCreated)"
$subject = "⚠ Rogue Account Disabled: $($user.SamAccountName)"

Send-MailMessage -From "CorporateEmail@corp.com" -To "TechStaff@corp.com" -Subject $subject -Body $body -SmtpServer "smtp.office365.com" -Credential $EmailCredential -UseSsl -Port 587

Write-Host "Disabled rogue account: $newUser"
