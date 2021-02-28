<#
.SYNOPSIS
    Get information from uninstall keys in Registry
.DESCRIPTION
    Extract DisplaName, DisplayVersion, Publisher, InstallDate, UninstallString and IconPath from each Uninstall subkeys in Registry
.EXAMPLE
    PS C:\> Get-InstalledSoftware -Name "*Reader*" -Company "Adobe*"
    Find software whith DisplaName, that contains Reader and Company name starts with Adobe
.EXAMPLE
    PS C:\> Get-InstalledSoftware -InstalledAfter 2018/01/01 -InstalledBefore 2019/01/01
    Find Software, which was installed after 2018/01/01 before 2019/01/01
.INPUTS
    Name - Software product name
    Company - Software publisher
    InstallBefore - Select software which was installed before InstallBefore date
    InstallAfter - Select software which was installed after InstallAfter date
    ShowHidden - Include also software with missed DisplayName, UninstallString and if SystemComponent was set
.OUTPUTS
    Object(s) include the next properties:
    - [string]Company
    - [string]Name
    - [string]Version
    - [DateTime]InstallDate(if was specified)
    - [string]UninstallString
    - [string]IconPath
.NOTES
    Andrey Vyporkhanyuk, 2019
#>
function Get-InstalledSoftware {
    [CmdletBinding()]
    Param (
        # Name of the software
        [Parameter(Position = 0, ValueFromPipeline = $true)]
        [string[]]$Name = '*',
        
        # Product Company
        [Parameter(Position = 1)][string[]]$Company = '*',
        
        # # Select software installed before specified date
        [Parameter(Position = 3)][DateTime]$InstallBefore,
        
        # # Select software installed after specified date
        [Parameter(Position = 4)][DateTime]$InstallAfter,

        # Select also hidden strings
        [switch]$ShowHidden )
    
    begin {
        $minDate = [DateTime]::MinValue
        $maxDate = [DateTime]::MaxValue

        $UninstallKeys = @( "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*", 
            "HKLM:\SOFTWARE\WOw6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
            "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*" ) 
    }
    
    process {
        $FoundItems = Get-ItemProperty  $UninstallKeys |
            Where-Object {
                (-not $_.SystemComponent -or $ShowHidden) -and ($_.DisplayName -or $ShowHidden) `
                    -and ($_.UninstallString -or $ShowHidden) -and ($_.DisplayName -like "$Name") `
                    -and ($_.Publisher -like "$Company") `
                    -and $(if ($InstallBefore) {
                        switch -Regex ($_.InstallDate) {
                            '^[0-9]' { $InstallDate = [int]$_ ; break }
                            default { $InstallDate = $maxDate }
                        } #switch
                        $InstallDate -le [int]"$($InstallBefore.Year) $($InstallBefore.Month) $($InstallBefore.Day)".Replace(" ", "") 
                    } # if
                    else { $true }) `
                    -and $(if ($InstallAfter) {
                        switch -Regex ($_.InstallDate) {
                            '^[0-9]' { $InstallDate = [int]$_; break }
                            default { $InstallDate = $minDate }
                        } #switch
                        $InstallDate -ge [int]$InstallAfter.ToString("yyyyMMdd") 
                    } # if
                    else { $true }) } |
            Select-Object -Property * -ExcludeProperty PS*
    
        foreach ($FoundItem in $FoundItems) {
            [PSCustomObject]@{
                'Company'         = $FoundItem.Publisher
                'Name'            = $FoundItem.DisplayName
                'Version'         = $FoundItem.DisplayVersion
                'InstallDate'     = $(
                    try { [DateTime]::ParseExact($FoundItem.InstallDate, "yyyyMMdd", $null) }
                    catch { $null } )
                'UninstallString' = $FoundItem.UninstallString
                'IconPath'        = $FoundItem.DisplayIcon
            }
        }
    }
    end { }
}