# --- Configurable Settings ---
$ServiceAccount = 'FPWiz'
$DaysUntilExpiration = 25
$CurrentDate = (Get-Date -Hour 0 -Minute 0 -Second 0)
$AlertRecipients = Get-Content "$PSScriptRoot\FPWIZ_Alert_Emails.txt"
$EmailTo = $AlertRecipients -join ';'
$ScriptToMonitorPWChange = 'C:\Scripts\Email_Alert\Watch-PWChange.ps1'  # Update this path if needed

# --- Fetch password expiration info ---
$PasswordLastSet = (Get-ADUser -Identity $ServiceAccount -Properties PasswordLastSet).PasswordLastSet
$PasswordExpiryDate = $PasswordLastSet.AddDays($DaysUntilExpiration)

# --- Trigger alert if password expiration is within 8 days ---
if ($PasswordExpiryDate.AddDays(-8) -le $CurrentDate) {
    $DaysRemaining = ($PasswordExpiryDate - $CurrentDate).Days
    $ExpiryDateString = $PasswordExpiryDate.ToString("dddd, MMMM dd, yyyy")
    $MeetingSubject = "$ServiceAccount Password Expiry"
    $EmailSubject = "URGENT: $ServiceAccount Password Expires in $DaysRemaining Days"

    $EmailBody = @"
Service Account: $ServiceAccount is expiring soon!

Expiration Date:    $ExpiryDateString
Days Remaining:     $DaysRemaining

Please change the password and update dependent services accordingly.
"@

    # --- Outlook COM: Send Email + Create Calendar Event ---
    $Outlook = New-Object -ComObject Outlook.Application
    $Namespace = $Outlook.GetNamespace("MAPI")
    $Calendar = $Namespace.GetDefaultFolder(9) # 9 = olFolderCalendar
    $Items = $Calendar.Items
    $Items.IncludeRecurrences = $true
    $Items.Sort("[Start]")

    $MeetingStart = $PasswordExpiryDate.Date
    $MeetingEnd = $MeetingStart.AddHours(23).AddMinutes(59)

    # --- Check for existing meeting ---
    $ExistingMeeting = $Items | Where-Object {
        $_.Start -ge $MeetingStart -and
        $_.Start -lt $MeetingEnd -and
        $_.AllDayEvent -eq $true -and
        $_.Subject -eq $MeetingSubject
    }

    if (-not $ExistingMeeting) {
        $Meeting = $Outlook.CreateItem(1) # olAppointmentItem
        $Meeting.Start       = $MeetingStart
        $Meeting.End         = $MeetingEnd
        $Meeting.Subject     = $MeetingSubject
        $Meeting.AllDayEvent = $true
        $Meeting.BusyStatus  = 0 # Free
        $Meeting.ReminderSet = $true
        $Meeting.ReminderMinutesBeforeStart = 0
        $Meeting.Save()
    }

    # --- Send alert email ---
    $Mail = $Outlook.CreateItem(0)
    $Mail.BCC = $EmailTo
    $Mail.Subject = $EmailSubject
    $Mail.Body = $EmailBody
    $Mail.Send()

    # --- Cleanup COM objects ---
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Items)     | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Calendar)  | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Namespace) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Outlook)   | Out-Null

    # --- Register scheduled task to watch for password change ---
    if (Test-Path $ScriptToMonitorPWChange) {
        $TaskName = "$ServiceAccount-PWChangeMonitor"
        $TaskTime = (Get-Date $PasswordExpiryDate -Hour 9 -Minute 0 -Second 0)

        $Trigger = New-ScheduledTaskTrigger -Once -At $TaskTime
        $Action  = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-ExecutionPolicy Bypass -WindowStyle Hidden -File `"$ScriptToMonitorPWChange`""

        Register-ScheduledTask -TaskName $TaskName -Trigger $Trigger -Action $Action -Force | Out-Null
        Write-Host "Scheduled task created: $TaskName"
    } else {
        Write-Warning "Script not found: $ScriptToMonitorPWChange — scheduled task not created."
    }
}
