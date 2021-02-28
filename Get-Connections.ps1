<#
.SYNOPSIS
    Resolve Application network connections
.DESCRIPTION
    Script get application path or shortcut path, launch it and detects its network connections
.EXAMPLE
    PS C:\> Get-Connections -Path <FullPath>
.Parameter Path
    The Path to the application or shortcut
#>

[CmdletBinding()]
Param(
    [Parameter(Mandatory=$True)]
    [String]$Path
)
$SLEEP_TIMEOUT = 3 #seconds
$ProcessList = @()
$ProcessList += Start-Process -FilePath $Path -PassThru -WindowStyle Minimized | Select-Object -ExpandProperty Id
Start-Sleep -Seconds $SLEEP_TIMEOUT
$ProcessList += Get-WmiObject -Class Win32_Process -Filter "ParentProcessID=$ProcessList" | Select-Object -ExpandProperty ProcessId

$TestTime = 0
$MAX_TEST_TIME = 24
Do {
    ForEach ($ProcId in $ProcessList) {
        Get-NetTCPConnection -OwningProcess $ProcId -ErrorAction Ignore | Format-Table -Property RemoteAddress, RemotePort
    }
    Write-Host "wait for ${SLEEP_TIMEOUT} seconds... "
    Start-Sleep -Seconds $SLEEP_TIMEOUT
    $TestTime += 1
} While ($TestTime -lt $MAX_TEST_TIME)

PAUSE