<#
.SYNOPSIS
    Update installed Microsoft Office 365 to the current version
.DESCRIPTION
    Module uses built-in Microsoft Office 365 tool to update it to the latest version. It checks current
    configuration for changes: Product Release, Shared Computer Licensing and Platform. Module use timeout for
    waiting on update process to be completed. Variable $GET_VERSION_MAX_TIME controls time for getting 
    the latest Office version. Variables $WAIT_UPDATE_MAX_TIME controls time for update installation.
    IMPORTANT! Module has to be run before Office 365 installation.
.NOTES
    AUTHOR:             Andre.Rovik@cegal.com
    HISTORY:
        [YYYY-MM-DD]    1.0     initial by Andre.Rovik@cegal.com
        [YYYY-MM-DD]    1.1     updated by thor.magnus.lilleaas@cegal.com
        [2022-04-01]    1.2     updated by andrii.vyporkhaniuk@cegal.com
#>
Param()

# Max time to wait for getting new version info
$GET_VERSION_MAX_TIME = 60 # seconds

# Max time to wait for update process
$WAIT_UPDATE_MAX_TIME = 30 # minutes

function Write-InstallLog {
    [CmdletBinding()]
    param( [Parameter(Mandatory=$True)]$Message, [switch]$isError )

    if ($isError) {
        $Message = $Message.ToUpper()
        "$(Get-Date) ERROR: $Message".Replace("UserOutput:", '') | Out-File $LogFile -Append
        throw $Message
    }
    else {
        "$(Get-Date): $Message".Replace("UserOutput:", '') | Out-File $LogFile -Append
        Write-Output $Message }
}

$LogFile = "$env:WinDir\Logs\Microsoft-Office365-Update_Script.txt"

# Get current Office 365 information
$ConfigKey = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
$OfficeProps = Get-ItemProperty -Path $ConfigKey |
    Select-Object -Property 'ProductReleaseIds', 'SharedComputerLicensing', 'Platform', 'VersionToReport'

$WelcomeText = @"
Office 365 edition: $($OfficeProps.ProductReleaseIds)
Shared Computer Licensing: $($OfficeProps.SharedComputerLicensing)
Platform: $($OfficeProps.Platform)
Current installed version: $($OfficeProps.VersionToReport)

***** Starting Office 365 update *****
"@

Write-InstallLog -Message $WelcomeText

# Start Office 365 updater utility
$StartParams = @{
    'FilePath' = "$env:CommonProgramFiles\Microsoft Shared\ClickToRun\OfficeC2RClient.exe"
    'ArgumentList' = '/update USER displaylevel=False'
    'WindowStyle' = 'Hidden'
    'Wait' = $false
}
Write-InstallLog -Message "Run: $($StartParams.FilePath) $($StartParams.ArgumentList)"
Start-Process @StartParams

# Wait to check if new version available
$counter = 0
$NewVersion = ""
$UpdateKey = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Updates"
do {
    Start-Sleep -Seconds 1
    $NewVersion = (Get-ItemProperty -Path $UpdateKey).UpdateToVersion
    if ($NewVersion) {
        Write-InstallLog -Message "UserOutput: New version is detected: $NewVersion"
        break
    }
} while ($counter++ -lt $GET_VERSION_MAX_TIME)

# Exit script if new version was not detected
if (-not $NewVersion) {
    Write-InstallLog -Message "UserOutput: No new version detected"
    return
}

# Wait for Office 365 update
Write-InstallLog "Office 365 update in progress, please wait... "
$counter = 0
do {
    Write-HOst $counter
    Start-Sleep -Seconds 60
    $NewOfficeProps = Get-ItemProperty -Path $ConfigKey |
    Select-Object -Property 'ProductReleaseIds', 'SharedComputerLicensing', 'Platform', 'VersionToReport'
    if ($NewVersion -eq $NewOfficeProps.VersionToReport) {
        # Check Office 365 configuration changes:
        Write-InstallLog "Office 365 edition is OK: $($OfficeProps.ProductReleaseIds -eq $NewOfficeProps.ProductReleaseIds)"
        Write-InstallLog "Shared Computer Licensing is OK: $($OfficeProps.SharedComputerLicensing -eq $NewOfficeProps.SharedComputerLicensing)"
        Write-InstallLog "Office 365 Platform is OK: $($OfficeProps.Platform -eq $NewOfficeProps.Platform)"
        Write-InstallLog "Current version is $($NewOfficeProps.VersionToReport)"

        Write-InstallLog -Message "UserOutput: Success."
        return
    }
} while ($counter++ -lt $WAIT_UPDATE_MAX_TIME)

if ($counter -ge $WAIT_UPDATE_MAX_TIME) {
    Write-InstallLog -Message "Update failed: Max wait time was reached." -isError
}