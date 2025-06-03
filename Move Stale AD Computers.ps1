<#
.SYNOPSIS
Identifies stale computer accounts in AD and optionally moves them to a target OU.

.DESCRIPTION
This script searches AD for enabled computers that haven't logged on in X days,
backs up their info, and optionally moves them. Errors are logged to file only.

.PARAMETER SearchOU
The Distinguished Name of the OU to search for computer accounts.

.PARAMETER TargetPath
The Distinguished Name of the OU to move stale computers to.

.PARAMETER DaysInactive
The number of days a computer must be inactive to be considered stale.

.PARAMETER LogPath
The directory to save CSV logs and error logs.

.PARAMETER DryRun
If set, displays what would be moved without actually performing the move.

.PARAMETER PromptUser
If set, prompts for confirmation before moving each computer.
#>

Param(
    [string]$SearchOU = 'OU=Computers,DC=domain,DC=com',
    [string]$TargetPath = 'OU=StaleComputers,DC=domain,DC=com',
    [int]$DaysInactive = 45,
    [string]$LogPath = 'C:\Temp',
    [switch]$DryRun,
    [switch]$PromptUser
)

# Create log folder if it doesn't exist
if (-not (Test-Path $LogPath)) {
    New-Item -ItemType Directory -Path $LogPath -Force | Out-Null
}

# Log setup
$LogDate = Get-Date -Format 'MM-dd-yyyy'
$StaleListPath = Join-Path $LogPath "StaleComputerList-$LogDate.csv"
$MovedListPath = Join-Path $LogPath "Moved_Computers-$LogDate.csv"
$ErrorLogPath = Join-Path $LogPath "Stale_Computers_Errors-$LogDate.log"
$BackupListPath = Join-Path $LogPath "Backup_StaleComputers-$LogDate.csv"

# Verify target OU exists
try {
    if (-not (Get-ADOrganizationalUnit -Filter "DistinguishedName -eq '$TargetPath'" -ErrorAction Stop)) {
        throw "Target OU '$TargetPath' does not exist."
    }
} catch {
    Add-Content -Path $ErrorLogPath -Value "FATAL: $_"
    exit 1
}

# Calculate inactive threshold
$InactiveThreshold = (Get-Date).AddDays(-$DaysInactive)

# Get stale computers
$StaleComputers = Get-ADComputer -SearchBase $SearchOU -Filter {
    LastLogonTimestamp -lt $InactiveThreshold -and Enabled -eq $true
} -Properties LastLogonTimestamp, DistinguishedName

# Export backup of full stale list
$StaleComputers | Export-Csv -Path $BackupListPath -NoTypeInformation

# Export readable stale list
$StaleComputers | Select-Object Name, Enabled,
    @{Name='LastLogonDate'; Expression = {[DateTime]::FromFileTime($_.LastLogonTimestamp)}},
    @{Name='Original OU'; Expression = {
        if ($_.DistinguishedName -match '(OU=.*?)(,DC=|$)') { $matches[1] } else { 'N/A' }
    }} | Export-Csv -Path $StaleListPath -NoTypeInformation

# Begin processing
foreach ($computer in $StaleComputers) {
    $compName = $computer.Name

    try {
        # Skip if no timestamp
        if (-not $computer.LastLogonTimestamp) {
            Add-Content -Path $ErrorLogPath -Value "$compName : Missing LastLogonTimestamp. Skipped."
            continue
        }

        $lastLogon = [DateTime]::FromFileTime($computer.LastLogonTimestamp)
        if ($lastLogon -lt (Get-Date).AddYears(-10)) {
            Add-Content -Path $ErrorLogPath -Value "$compName : Last logon over 10 years ago. Skipped."
            continue
        }

        # Prompt user if enabled
        if ($PromptUser) {
            $confirmation = Read-Host "Move $compName to $TargetPath? (Y/N)"
            if ($confirmation -notin @('Y', 'y')) {
                Add-Content -Path $ErrorLogPath -Value "$compName : Skipped by user input."
                continue
            }
        }

        # Dry run display
        if ($DryRun) {
            Add-Content -Path $ErrorLogPath -Value "[DryRun] $compName would be moved to $TargetPath"
            continue
        }

        # Get object and move
        $CompObj = Get-ADComputer -Identity $compName -Properties DistinguishedName
        $MovedObj = $CompObj | Move-ADObject -TargetPath $TargetPath -PassThru -ErrorAction Stop

        $MovedObj | Select-Object Name, Enabled,
            @{Name = "New OU"; Expression = {
                if ($_.DistinguishedName -match '(OU=.*?)(,DC=|$)') { $matches[1] } else { "N/A" }
            }} | Export-Csv -Path $MovedListPath -NoTypeInformation -Append
    }
    catch {
        Add-Content -Path $ErrorLogPath -Value "$compName : $($_.Exception.Message)"
    }
}
